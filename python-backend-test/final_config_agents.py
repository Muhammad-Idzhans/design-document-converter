"""
final_config_agents.py — Final Configuration Document generator agents and prompts.

This module handles:
1. Azure Resource Proxy endpoints for browsing resources via Entra ID tokens
2. (Future) AI prompts specific to final config document generation
3. (Future) Background processing pipeline for final config
"""

import json
import os
import traceback
from pathlib import Path
import requests
from typing import List, Dict, Any, Optional
import re
import json

from config import (
    AGENT_OPENAI_ENDPOINT, AGENT_OPENAI_KEY, FINAL_CONFIG_AGENT_ID,
    UPLOAD_DIR, WRITER_DEPLOYMENT
)
from task_manager import update_task, get_task, upload_task_to_blob
from final_config_docx import build_final_config_docx

ARM_BASE_URL = "https://management.azure.com"
API_VERSION_SUBS = "2022-12-01"          # subscriptions & resource groups
API_VERSION_RESOURCES = "2021-04-01"     # resources list

def _arm_get(url: str, token: str) -> List[Dict[str, Any]]:
    """Call an ARM REST endpoint and return JSON, following pagination."""
    # Ensure token format is correct (sometimes frontend sends 'Bearer ', sometimes just the token)
    if not token.startswith("Bearer "):
        token = f"Bearer {token}"
        
    headers = {"Authorization": token}
    items = []
    
    while url:
        resp = requests.get(url, headers=headers, timeout=60)
        
        # If unauthorized, raise an exception that FastAPI can catch
        if resp.status_code == 401 or resp.status_code == 403:
            raise ValueError(f"Azure API Authentication Error: {resp.status_code}. Token may be invalid or expired.")
            
        if resp.status_code != 200:
            raise RuntimeError(f"Azure API error ({resp.status_code}): {resp.text[:500]}")
            
        data = resp.json()
        items.extend(data.get("value", []))
        url = data.get("nextLink")  # ARM paginates large results
        
    return items

def get_azure_subscriptions(token: str) -> List[Dict[str, Any]]:
    """List all subscriptions the user has access to."""
    url = f"{ARM_BASE_URL}/subscriptions?api-version={API_VERSION_SUBS}"
    subs = _arm_get(url, token)
    return [
        {
            "subscriptionId": s.get("subscriptionId"),
            "displayName": s.get("displayName"),
            "state": s.get("state"),
        }
        for s in subs
    ]

def get_azure_resource_groups(token: str, subscription_id: str) -> List[Dict[str, Any]]:
    """List all resource groups inside a subscription."""
    url = f"{ARM_BASE_URL}/subscriptions/{subscription_id}/resourcegroups?api-version={API_VERSION_SUBS}"
    groups = _arm_get(url, token)
    return [
        {
            "name": g.get("name"),
            "location": g.get("location"),
            "managedBy": g.get("managedBy"),
        }
        for g in groups
    ]

def get_azure_resources(token: str, subscription_id: str, rg_name: str) -> List[Dict[str, Any]]:
    """List all resources inside a subscription + resource group."""
    url = (
        f"{ARM_BASE_URL}/subscriptions/{subscription_id}"
        f"/resourceGroups/{rg_name}/resources"
        f"?api-version={API_VERSION_RESOURCES}"
    )
    resources = _arm_get(url, token)
    return [
        {
            "id": r.get("id"),
            "name": r.get("name"),
            "type": r.get("type"),
            "location": r.get("location"),
            "sku": r.get("sku"),
            "kind": r.get("kind"),
        }
        for r in resources
    ]

def _arm_get_single(url: str, token: str) -> Dict[str, Any]:
    """Call an ARM REST endpoint and return single JSON object."""
    if not token.startswith("Bearer "):
        token = f"Bearer {token}"
        
    headers = {"Authorization": token}
    resp = requests.get(url, headers=headers, timeout=60)
    
    if resp.status_code in [401, 403]:
        raise ValueError(f"Azure API Authentication Error: {resp.status_code}. Token may be invalid or expired.")
        
    if resp.status_code != 200:
        raise RuntimeError(f"Azure API error ({resp.status_code}): {resp.text[:500]}")
        
    return resp.json()



def _parse_resource_id(resource_id: str) -> Dict[str, str]:
    """Extract subscription ID, provider namespace and full resource type from an ARM resource ID.

    Example:
      /subscriptions/xxx/resourceGroups/rg/providers/Microsoft.CognitiveServices/accounts/a1/projects/p1
      -> namespace = Microsoft.CognitiveServices, resource_type = accounts/projects
    """
    parts = [p for p in resource_id.strip("/").split("/") if p]
    lower = [p.lower() for p in parts]

    sub_id = ""
    if "subscriptions" in lower:
        idx = lower.index("subscriptions")
        if idx + 1 < len(parts):
            sub_id = parts[idx + 1]

    if "providers" not in lower:
        return {"subscription_id": sub_id, "namespace": "", "resource_type": ""}

    # Use the LAST 'providers' segment (handles extension resources)
    prov_idx = len(lower) - 1 - lower[::-1].index("providers")
    after = parts[prov_idx + 1:]  # [namespace, type1, name1, type2, name2, ...]
    namespace = after[0] if after else ""
    type_segments = after[1::2]   # type1, type2, ... (skips the names)

    return {
        "subscription_id": sub_id,
        "namespace": namespace,
        "resource_type": "/".join(type_segments),
    }


def _get_api_version(token: str, subscription_id: str, namespace: str,
                     resource_type: str, cache: Dict[str, str]) -> str:
    """Look up the latest supported api-version for a resource type via the ARM providers endpoint.

    Prefers the latest stable (non-preview) version. Falls back to the latest preview,
    then to API_VERSION_RESOURCES if the lookup fails. Results are cached per type.
    """
    cache_key = f"{namespace}/{resource_type}".lower()
    if cache_key in cache:
        return cache[cache_key]

    version = API_VERSION_RESOURCES  # safe fallback
    try:
        url = (
            f"{ARM_BASE_URL}/subscriptions/{subscription_id}"
            f"/providers/{namespace}?api-version={API_VERSION_RESOURCES}"
        )
        provider = _arm_get_single(url, token)

        for rt in provider.get("resourceTypes", []):
            if rt.get("resourceType", "").lower() == resource_type.lower():
                versions = sorted(rt.get("apiVersions", []), reverse=True)
                stable = [v for v in versions if "preview" not in v.lower()]
                if stable:
                    version = stable[0]
                elif versions:
                    version = versions[0]
                break
        else:
            print(f"[api-version] Type '{resource_type}' not found under {namespace}; using fallback {version}")
    except Exception as e:
        print(f"[api-version] Lookup failed for {namespace}/{resource_type}: {e}; using fallback {version}")

    cache[cache_key] = version
    print(f"[api-version] {namespace}/{resource_type} -> {version}")
    return version

def _norm_location(value: str) -> str:
    return (value or "").lower().replace(" ", "")


def _get_location_map(token: str, subscription_id: str, cache: dict) -> dict:
    """Map ARM location names ('swedencentral') to display names ('Sweden Central'). Cached per subscription."""
    if subscription_id in cache:
        return cache.get(subscription_id)
    mapping = {}
    try:
        url = f"{ARM_BASE_URL}/subscriptions/{subscription_id}/locations?api-version={API_VERSION_SUBS}"
        for loc in _arm_get(url, token):
            display = loc.get("displayName")
            if display:
                mapping[_norm_location(loc.get("name", ""))] = display
                mapping[_norm_location(display)] = display
    except Exception as e:
        print(f"[locations] Lookup failed for {subscription_id}: {e}")
    cache[subscription_id] = mapping
    return mapping

def fetch_resource_details(resource_ids: List[str], token: str) -> Dict[str, List[Dict[str, Any]]]:
    """Fetch detailed JSON for each Azure Resource ID using the correct api-version per type.

    Returns:
        {
          "succeeded": [ <ARM resource JSON>, ... ],
          "failed":    [ {"id", "type", "api_version", "reason"}, ... ]
        }
    Only 'succeeded' should ever be sent to the AI.
    """
    succeeded: List[Dict[str, Any]] = []
    failed: List[Dict[str, Any]] = []
    version_cache: Dict[str, str] = {}
    location_cache: dict = {}

    for res_id in resource_ids:
        if not res_id.startswith("/"):
            res_id = f"/{res_id}"

        info = _parse_resource_id(res_id)
        if not info["subscription_id"] or not info["namespace"] or not info["resource_type"]:
            failed.append({
                "id": res_id,
                "type": "unknown",
                "api_version": None,
                "reason": "Could not parse resource ID",
            })
            continue

        full_type = f"{info['namespace']}/{info['resource_type']}"
        api_version = _get_api_version(
            token, info["subscription_id"], info["namespace"], info["resource_type"], version_cache
        )

        url = f"{ARM_BASE_URL}{res_id}?api-version={api_version}"
        try:
            data = _arm_get_single(url, token)

            # Phase 4: extra data captured while the token is fresh (keys starting with _ are internal)
            data["_resource_group"] = _resource_group_from_id(res_id)
            loc_map = _get_location_map(token, info["subscription_id"], location_cache)
            raw_loc = data.get("location") or ""
            data["_location_display"] = loc_map.get(_norm_location(raw_loc), raw_loc)

            if full_type.lower() == "microsoft.web/sites":
                try:
                    cfg_url = f"{ARM_BASE_URL}{res_id}/config/web?api-version={api_version}"
                    data["_siteConfig"] = _arm_get_single(cfg_url, token).get("properties") or {}
                except Exception as e:
                    print(f"[fetch_resource_details] config/web skipped for {res_id.split('/')[-1]}: {e}")

            succeeded.append(data)
            print(f"[fetch_resource_details] OK   {full_type} ({api_version}) {res_id.split('/')[-1]}")
        except Exception as e:
            print(f"[fetch_resource_details] FAIL {full_type} ({api_version}) {res_id}: {e}")
            failed.append({
                "id": res_id,
                "type": full_type,
                "api_version": api_version,
                "reason": str(e)[:500],
            })

    print(f"[fetch_resource_details] Done: {len(succeeded)} succeeded, {len(failed)} failed")
    return {"succeeded": succeeded, "failed": failed}


# def generate_final_config_markdown(task_id: str, resources_data: List[Dict[str, Any]]) -> str:
#     task = get_task(task_id)
#     project_title = task.get("projectTitle", "Azure Configuration")
#     client_name = task.get("clientName", "Client")
    
#     base_endpoint = AGENT_OPENAI_ENDPOINT.split("/openai")[0].rstrip("/")
#     url = f"{base_endpoint}/openai/v1/chat/completions"
#     headers = {
#         "api-key": AGENT_OPENAI_KEY,
#         "Authorization": f"Bearer {AGENT_OPENAI_KEY}",
#         "Content-Type": "application/json"
#     }

#     writer_system_prompt = """
#         You are a Cloud Infrastructure Architect. Your task is to analyze raw Azure Resource Manager (ARM) JSON configurations and generate a formal 'Final Configuration Document' in Markdown.
        
#         CRITICAL INSTRUCTIONS:
#         1. GROUND TRUTH: You must ONLY use the provided JSON data. Do not hallucinate settings, regions, or configurations that are not explicitly in the JSON.
#         2. NO WEB SEARCH: You must NOT use web search. Rely entirely on the provided ARM JSON.
#         3. NO PLACEHOLDERS: Do not use placeholders like [CLIENT_NAME]. The client is explicitly provided.
#         4. STRUCTURE: Group the resources logically (e.g., Compute, Networking, Storage). For each resource, explain its purpose and present its key configurations in a Markdown table.
#         5. TABLES: Ensure all tables have proper markdown syntax with leading and trailing pipes `|`. Leave a blank line before and after each table.
#         6. NO CODE BLOCKS: Do not wrap your final output in a ```markdown code block. Just return the raw text.
#     """
    
#     use_agent = False
#     openai_client = None
#     agent_name = ""
#     agent_version = "1"

#     if FINAL_CONFIG_AGENT_ID:
#         try:
#             from azure.identity import DefaultAzureCredential
#             from azure.ai.projects import AIProjectClient
#             project_client = AIProjectClient(
#                 endpoint=AGENT_OPENAI_ENDPOINT,
#                 credential=DefaultAzureCredential()
#             )
#             openai_client = project_client.get_openai_client()
#             agent_parts = FINAL_CONFIG_AGENT_ID.split(":")
#             agent_name = agent_parts[0]
#             agent_version = agent_parts[1] if len(agent_parts) > 1 else "1"
#             use_agent = True
#         except Exception as e:
#             print(f"Failed to intialize Azure AI Projects Client for Final Config: {e}")

#     user_prompt = (
#         f"Generate the Final Configuration Document for {client_name}'s project: {project_title}.\n\n"
#         f"Raw Azure ARM JSON Data:\n{json.dumps(resources_data, indent=2)}"
#     )

#     drafted_text = None

#     if use_agent and openai_client:
#         try:
#             response = openai_client.responses.create(
#                 input=[
#                     {"type": "message", "role": "system", "content": writer_system_prompt},
#                     {"type": "message", "role": "user", "content": user_prompt}
#                 ],
#                 extra_body={
#                     "agent_reference": {
#                         "name": agent_name,
#                         "version": agent_version,
#                         "type": "agent_reference"
#                     }
#                 }
#             )
#             drafted_text = response.output_text
            
#             try:
#                 usage = response.usage
#                 if task and "cost_metrics" in task:
#                     prompt_tokens = getattr(usage, 'input_tokens', 0) or getattr(usage, 'prompt_tokens', 0)
#                     completion_tokens = getattr(usage, 'output_tokens', 0) or getattr(usage, 'completion_tokens', 0)
#                     task["cost_metrics"]["llm_tokens_prompt"] += prompt_tokens
#                     task["cost_metrics"]["llm_tokens_completion"] += completion_tokens
#                     update_task(task_id, {"cost_metrics": task["cost_metrics"]})
#             except Exception as e:
#                 print(f"[generate_final_config] Agent usage token extraction failed: {e}")
#         except Exception as e:
#             print(f"[generate_final_config] Agent drafting failed: {e}")

#     if not use_agent or not drafted_text:
#         payload = {
#             "model": WRITER_DEPLOYMENT,
#             "messages": [
#                 {"role": "system", "content": writer_system_prompt},
#                 {"role": "user", "content": user_prompt}
#             ],
#             "temperature": 0.1,
#             "max_tokens": 16384
#         }
#         try:
#             response = requests.post(url, headers=headers, json=payload, timeout=240)
#             if response.status_code == 200:
#                 data = response.json()
#                 drafted_text = data["choices"][0]["message"]["content"].strip()
                
#                 usage = data.get("usage", {})
#                 if task and "cost_metrics" in task:
#                     task["cost_metrics"]["llm_tokens_prompt"] += usage.get("prompt_tokens", 0)
#                     task["cost_metrics"]["llm_tokens_completion"] += usage.get("completion_tokens", 0)
#                     update_task(task_id, {"cost_metrics": task["cost_metrics"]})
#         except Exception as e:
#             print(f"[generate_final_config] Fallback LLM call failed: {e}")

#     if drafted_text:
#         drafted_text = drafted_text.strip()
#         if drafted_text.startswith("```"):
#             drafted_text = drafted_text.split("\n", 1)[-1]
#         if drafted_text.endswith("```"):
#             drafted_text = drafted_text.rsplit("\n", 1)[0]
#         return drafted_text.strip()
    
#     return "*(Error generating final configuration document)*"

# =============================================================================
# Phase 3 — Code-built document structure (Broadfield pattern)
# Code writes headings, intro sentences, tables and captions.
# AI writes ONLY the short Overview paragraph per service.
# =============================================================================

CATEGORY_ORDER = ["Compute", "Networking", "Databases", "Storage",
                  "AI Services", "Monitoring", "Security", "Other"]

# ARM type (lowercase) -> (Section heading, Category). Order here = order within a category.
SERVICE_CATALOG = {
    # Compute
    "microsoft.compute/virtualmachines": ("Azure Virtual Machine", "Compute"),
    "microsoft.web/serverfarms": ("Azure App Service Plan", "Compute"),
    "microsoft.web/sites": ("Azure App Service", "Compute"),
    "microsoft.app/containerapps": ("Azure Container Apps", "Compute"),
    "microsoft.containerservice/managedclusters": ("Azure Kubernetes Service", "Compute"),
    "microsoft.containerregistry/registries": ("Azure Container Registry", "Compute"),
    "microsoft.logic/workflows": ("Azure Logic Apps", "Compute"),
    # Networking
    "microsoft.network/virtualnetworks": ("Azure Virtual Network", "Networking"),
    "microsoft.network/networksecuritygroups": ("Network Security Group", "Networking"),
    "microsoft.network/azurefirewalls": ("Azure Firewall", "Networking"),
    "microsoft.network/firewallpolicies": ("Azure Firewall Policy", "Networking"),
    "microsoft.network/publicipaddresses": ("Public IP Address", "Networking"),
    "microsoft.network/applicationgateways": ("Azure Application Gateway", "Networking"),
    "microsoft.network/loadbalancers": ("Azure Load Balancer", "Networking"),
    "microsoft.network/virtualnetworkgateways": ("Azure Virtual Network Gateway", "Networking"),
    "microsoft.network/bastionhosts": ("Azure Bastion", "Networking"),
    "microsoft.network/routetables": ("Route Table", "Networking"),
    "microsoft.network/privateendpoints": ("Private Endpoint", "Networking"),
    "microsoft.cdn/profiles": ("Azure Front Door / CDN Profile", "Networking"),
    # Databases
    "microsoft.documentdb/databaseaccounts": ("Azure Cosmos DB", "Databases"),
    "microsoft.sql/servers": ("Azure SQL Server", "Databases"),
    "microsoft.sql/servers/databases": ("Azure SQL Database", "Databases"),
    "microsoft.dbformysql/flexibleservers": ("Azure Database for MySQL", "Databases"),
    "microsoft.dbforpostgresql/flexibleservers": ("Azure Database for PostgreSQL", "Databases"),
    "microsoft.cache/redis": ("Azure Cache for Redis", "Databases"),
    # Storage
    "microsoft.storage/storageaccounts": ("Azure Storage Account", "Storage"),
    # AI Services
    "microsoft.cognitiveservices/accounts": ("Azure AI Services", "AI Services"),
    "microsoft.cognitiveservices/accounts/projects": ("Azure AI Foundry Project", "AI Services"),
    "microsoft.search/searchservices": ("Azure AI Search", "AI Services"),
    "microsoft.machinelearningservices/workspaces": ("Azure Machine Learning Workspace", "AI Services"),
    # Monitoring
    "microsoft.operationalinsights/workspaces": ("Log Analytics Workspace", "Monitoring"),
    "microsoft.insights/components": ("Application Insights", "Monitoring"),
    # Security
    "microsoft.keyvault/vaults": ("Azure Key Vault", "Security"),
    "microsoft.recoveryservices/vaults": ("Recovery Services Vault", "Security"),
}

OVERVIEW_SYSTEM_PROMPT = """You write the 'Overview' paragraph for ONE Azure service in a Final Configuration Document prepared by Enfrasys Solutions.

Rules:
- Write ONE paragraph of 2 to 3 sentences. Plain text only.
- First sentence(s): a standard, factual definition of the Azure service.
- Last sentence: state how the service is present in the client's environment, using ONLY facts from the resource list (number of resources, names, region, SKU or kind).
- Do NOT invent purposes, workloads, security settings, deployment status, or anything not in the resource list.
- No headings, no bullet points, no tables, no Markdown formatting.
- Always refer to the client by the exact client name provided (for example "Contoso Ltd's environment"). Never write "the client".
- Write Azure regions exactly as given in the resource list.
"""


def _md_cell(value) -> str:
    """Make a value safe for one Markdown table cell (pipes/newlines break Pandoc tables)."""
    if value is None or value == "" or value == [] or value == {}:
        return "-"
    text = str(value).replace("\r", " ").replace("\n", " ").replace("|", "\\|").strip()
    return text or "-"


def _resource_group_from_id(resource_id: str) -> str:
    parts = resource_id.split("/")
    lower = [p.lower() for p in parts]
    if "resourcegroups" in lower:
        i = lower.index("resourcegroups")
        if i + 1 < len(parts):
            return parts[i + 1]
    return ""


def _display_name(resource: Dict[str, Any]) -> str:
    # Child resources come back as "parent/child" — show only the child name
    return (resource.get("name") or "").split("/")[-1]


# def _default_config_rows(resource: Dict[str, Any]) -> List:
#     """Safe, common fields only. Phase 4 adds richer per-type rows."""
#     props = resource.get("properties") or {}
#     sku = resource.get("sku") or {}
#     identity = resource.get("identity") or {}
#     tags = resource.get("tags") or {}

#     rows = [
#         ("Resource Name", _display_name(resource)),
#         ("Resource Group", _resource_group_from_id(resource.get("id", ""))),
#         ("Location", resource.get("location")),
#     ]
#     if sku.get("name"):
#         sku_text = sku["name"]
#         if sku.get("tier") and sku["tier"] != sku["name"]:
#             sku_text = f"{sku_text} ({sku['tier']})"
#         rows.append(("SKU", sku_text))
#     if resource.get("kind"):
#         rows.append(("Kind", resource["kind"]))
#     if "publicNetworkAccess" in props:
#         rows.append(("Public Network Access", props["publicNetworkAccess"]))
#     tls = props.get("minimumTlsVersion") or props.get("minimalTlsVersion")
#     if tls:
#         rows.append(("Minimum TLS Version", tls))
#     rows.append(("Managed Identity", identity.get("type") or "None"))
#     rows.append(("Tags", ", ".join(f"{k}: {v}" for k, v in tags.items()) if tags else "None"))
#     return rows

# -----------------------------
# Phase 4 — Blade-style configuration tables (one small table per portal section)
# Code picks every row. Rows without a value are skipped; empty tables are skipped.
# -----------------------------
def _loc(r: dict):
    return r.get("_location_display") or r.get("location")


def _rg(r: dict):
    return r.get("_resource_group") or _resource_group_from_id(r.get("id", ""))


def _yes(value):
    """True/False -> Enabled/Disabled. None stays None (row skipped)."""
    if value is None:
        return None
    return "Enabled" if value else "Disabled"


def _neg(value):
    """For ARM 'disable...' flags: disableLocalAuth=True -> Disabled."""
    if value is None:
        return None
    return "Disabled" if value else "Enabled"


def _tls(value):
    """'TLS1_2', 'Tls12', '1.2' -> 'TLS 1.2'."""
    if not value:
        return None
    digits = str(value).lower().replace("tls", "").replace("_", "").replace(".", "").strip()
    if digits.isdigit() and len(digits) >= 2:
        return f"TLS {digits[0]}.{digits[1:]}"
    return value


def _sku_text(r: dict):
    sku = r.get("sku") or {}
    name, tier = sku.get("name"), sku.get("tier")
    if not name:
        return None
    return f"{name} ({tier})" if tier and tier != name else name


def _identity_text(r: dict) -> str:
    ident = r.get("identity") or {}
    kind = ident.get("type")
    if not kind or kind.lower() == "none":
        return "None"
    user_ids = ident.get("userAssignedIdentities") or {}
    return f"{kind} ({len(user_ids)} user-assigned)" if user_ids else kind


def _tags_text(r: dict) -> str:
    tags = {k: v for k, v in (r.get("tags") or {}).items()
            if not str(k).lower().startswith("hidden-")}
    return ", ".join(f"{k}: {v}" for k, v in tags.items()) if tags else "None"


def _clean(rows: list) -> list:
    return [(k, v) for k, v in rows if v not in (None, "", [], {})]


def _blades(*pairs) -> list:
    """Each pair is (blade_name, rows). Returns only blades that still have rows after cleaning."""
    result = []
    for blade, rows in pairs:
        cleaned = _clean(rows)
        if cleaned:
            result.append((blade, cleaned))
    return result


def _base(r: dict) -> list:
    return [("Resource Group", _rg(r)), ("Location", _loc(r))]


def _ident_tags(r: dict) -> list:
    return [("Managed Identity", _identity_text(r)), ("Tags", _tags_text(r))]


# ---------- Per-type extractors ----------
def _blades_app_service_plan(r: dict) -> list:
    p = r.get("properties") or {}
    sku = r.get("sku") or {}
    return _blades(
        ("Overview", _base(r) + [
            ("Operating System", "Linux" if p.get("reserved") else "Windows"),
            ("Pricing Tier", _sku_text(r)),
            ("Status", p.get("status")),
            ("Apps Hosted", p.get("numberOfSites")),
        ] + _ident_tags(r)),
        ("Scale", [
            ("Instance Count", sku.get("capacity")),
            ("Maximum Workers", p.get("maximumNumberOfWorkers")),
            ("Zone Redundant", _yes(p.get("zoneRedundant"))),
            ("Elastic Scale", _yes(p.get("elasticScaleEnabled"))),
        ]),
    )


def _blades_app_service(r: dict) -> list:
    p = r.get("properties") or {}
    cfg = r.get("_siteConfig") or {}

    runtime = cfg.get("linuxFxVersion") or cfg.get("windowsFxVersion")
    if runtime:
        runtime = runtime.replace("|", " ")
    elif cfg.get("netFrameworkVersion"):
        runtime = f".NET {cfg.get('netFrameworkVersion')}"

    cert_enabled = p.get("clientCertEnabled")
    if cert_enabled is None:
        client_cert = None
    else:
        client_cert = f"Enabled ({p.get('clientCertMode')})" if cert_enabled else "Disabled"

    plan = (p.get("serverFarmId") or "").split("/")[-1] or None
    os_name = "Linux" if "linux" in (r.get("kind") or "").lower() else "Windows"

    return _blades(
        ("Overview", _base(r) + [
            ("App Service Plan", plan),
            ("Operating System", os_name),
            ("State", p.get("state")),
            ("Default Domain", p.get("defaultHostName")),
        ] + _ident_tags(r)),
        ("Configuration", [
            ("Runtime Stack", runtime),
            ("Always On", _yes(cfg.get("alwaysOn"))),
            ("HTTP 2.0", _yes(cfg.get("http20Enabled"))),
            ("FTP State", cfg.get("ftpsState")),
        ]),
        ("Networking", [
            ("Public Network Access", p.get("publicNetworkAccess")),
            ("HTTPS Only", _yes(p.get("httpsOnly"))),
            ("Minimum TLS Version", _tls(cfg.get("minTlsVersion"))),
            ("Client Certificate", client_cert),
        ]),
    )


def _blades_storage(r: dict) -> list:
    p = r.get("properties") or {}
    acls = p.get("networkAcls") or {}
    enc = p.get("encryption") or {}
    key_source = {
        "microsoft.storage": "Microsoft-managed keys",
        "microsoft.keyvault": "Customer-managed keys",
    }.get((enc.get("keySource") or "").lower(), enc.get("keySource"))

    return _blades(
        ("Overview", _base(r) + [
            ("Performance / Redundancy", _sku_text(r)),
            ("Kind", r.get("kind")),
            ("Access Tier", p.get("accessTier")),
            ("Hierarchical Namespace", _yes(p.get("isHnsEnabled"))),
        ] + _ident_tags(r)),
        ("Security", [
            ("Secure Transfer (HTTPS Only)", _yes(p.get("supportsHttpsTrafficOnly"))),
            ("Minimum TLS Version", _tls(p.get("minimumTlsVersion"))),
            ("Blob Anonymous Access", _yes(p.get("allowBlobPublicAccess"))),
            ("Shared Key Access", _yes(p.get("allowSharedKeyAccess"))),
            ("Encryption Key Source", key_source),
            ("Infrastructure Encryption", _yes(enc.get("requireInfrastructureEncryption"))),
        ]),
        ("Networking", [
            ("Public Network Access", p.get("publicNetworkAccess")),
            ("Network Default Action", acls.get("defaultAction")),
            ("Network Bypass", acls.get("bypass")),
        ]),
    )


def _blades_cosmos(r: dict) -> list:
    p = r.get("properties") or {}
    caps = {(c.get("name") or "").lower() for c in (p.get("capabilities") or [])}
    kind = (r.get("kind") or "").lower()

    if kind == "mongodb" or "enablemongo" in caps:
        api = "MongoDB"
    elif "enablecassandra" in caps:
        api = "Cassandra"
    elif "enabletable" in caps:
        api = "Table"
    elif "enablegremlin" in caps:
        api = "Gremlin"
    else:
        api = "NoSQL"

    capacity = p.get("capacityMode") or ("Serverless" if "enableserverless" in caps else "Provisioned Throughput")
    locations = sorted(p.get("locations") or [], key=lambda l: l.get("failoverPriority", 0))
    regions = ", ".join(l.get("locationName") for l in locations if l.get("locationName"))

    backup = p.get("backupPolicy") or {}
    periodic = backup.get("periodicModeProperties") or {}
    continuous = backup.get("continuousModeProperties") or {}
    interval = periodic.get("backupIntervalInMinutes")
    retention = periodic.get("backupRetentionIntervalInHours")

    return _blades(
        ("Overview", _base(r) + [
            ("API", api),
            ("Capacity Mode", capacity),
            ("Free Tier", _yes(p.get("enableFreeTier"))),
            ("Default Consistency Level", (p.get("consistencyPolicy") or {}).get("defaultConsistencyLevel")),
        ] + _ident_tags(r)),
        ("Replication", [
            ("Regions", regions),
            ("Automatic Failover", _yes(p.get("enableAutomaticFailover"))),
            ("Multi-region Writes", _yes(p.get("enableMultipleWriteLocations"))),
        ]),
        ("Networking", [
            ("Public Network Access", p.get("publicNetworkAccess")),
            ("Minimum TLS Version", _tls(p.get("minimalTlsVersion"))),
            ("Local Authentication (Keys)", _neg(p.get("disableLocalAuth"))),
        ]),
        ("Backup", [
            ("Backup Policy", backup.get("type")),
            ("Backup Interval", f"{interval} minutes" if interval else None),
            ("Backup Retention", f"{retention} hours" if retention else None),
            ("Backup Storage Redundancy", periodic.get("backupStorageRedundancy")),
            ("Continuous Backup Tier", continuous.get("tier")),
        ]),
    )


def _blades_ai_services(r: dict) -> list:
    p = r.get("properties") or {}
    acls = p.get("networkAcls") or {}
    return _blades(
        ("Overview", _base(r) + [
            ("Kind", r.get("kind")),
            ("Pricing Tier", _sku_text(r)),
            ("Endpoint", p.get("endpoint")),
            ("Custom Subdomain", p.get("customSubDomainName")),
        ] + _ident_tags(r)),
        ("Networking", [
            ("Public Network Access", p.get("publicNetworkAccess")),
            ("Network Default Action", acls.get("defaultAction")),
            ("Local Authentication (Keys)", _neg(p.get("disableLocalAuth"))),
        ]),
    )


def _blades_foundry_project(r: dict) -> list:
    p = r.get("properties") or {}
    full_name = r.get("name") or ""
    parent = full_name.split("/")[0] if "/" in full_name else None
    is_default = p.get("isDefault")
    return _blades(
        ("Overview", [("Parent AI Services Resource", parent)] + _base(r) + [
            ("Display Name", p.get("displayName")),
            ("Description", p.get("description")),
            ("Default Project", None if is_default is None else ("Yes" if is_default else "No")),
        ] + _ident_tags(r)),
    )


# ---------- Fallback for any type without an extractor ----------
_SENSITIVE_PARTS = ("guid", "principal", "tenant", "etag", "time", "date", "created", "modified",
                    "provisioningstate", "ipaddress", "outbound", "ftp", "user", "password",
                    "secret", "token", "key", "connection", "credential", "email", "thumbprint")


def _is_sensitive(key: str) -> bool:
    k = str(key).lower()
    if k == "id" or k.endswith("id") or k.endswith("ids"):
        return True
    return any(part in k for part in _SENSITIVE_PARTS)


def _humanize(key: str) -> str:
    out = ""
    for i, ch in enumerate(key):
        if ch.isupper() and i > 0 and not key[i - 1].isupper():
            out += " "
        out += ch
    return out[:1].upper() + out[1:]


def _flatten_props(d, prefix="", depth=0, out=None) -> list:
    """Flatten 'properties' into readable rows (max 2 levels), skipping sensitive/noisy fields."""
    if out is None:
        out = []
    for k, v in (d or {}).items():
        if _is_sensitive(k):
            continue
        label = f"{prefix} {_humanize(k)}".strip()
        if isinstance(v, dict):
            if depth < 1:
                _flatten_props(v, label, depth + 1, out)
        elif isinstance(v, list):
            if v and len(v) <= 5 and all(isinstance(x, (str, int, float, bool)) for x in v):
                out.append((label, ", ".join(str(x) for x in v)))
        elif isinstance(v, bool):
            out.append((label, "Enabled" if v else "Disabled"))
        elif v not in (None, ""):
            out.append((label, v))
    return out


def _blades_default(r: dict) -> list:
    props_rows = _flatten_props(r.get("properties") or {})
    return _blades(
        ("Overview", _base(r) + [
            ("SKU", _sku_text(r)),
            ("Kind", r.get("kind")),
        ] + _ident_tags(r)),
        ("Properties", props_rows[:25]),
    )


CONFIG_EXTRACTORS = {
    "microsoft.web/serverfarms": _blades_app_service_plan,
    "microsoft.web/sites": _blades_app_service,
    "microsoft.storage/storageaccounts": _blades_storage,
    "microsoft.documentdb/databaseaccounts": _blades_cosmos,
    "microsoft.cognitiveservices/accounts": _blades_ai_services,
    "microsoft.cognitiveservices/accounts/projects": _blades_foundry_project,
}


def get_config_blades(r: dict) -> list:
    """Returns a list of (blade_name, rows). Unknown types use the safe fallback."""
    extractor = CONFIG_EXTRACTORS.get((r.get("type") or "").lower(), _blades_default)
    try:
        return extractor(r)
    except Exception as e:
        print(f"[extractor] {r.get('type')} {_display_name(r)} failed: {e}; using fallback")
        return _blades_default(r)


def _config_table_md(rows: List[tuple]) -> str:
    lines = ["| Setting | Value |", "|---|---|"]
    lines += [f"| {_md_cell(k)} | {_md_cell(v)} |" for k, v in rows]
    return "\n".join(lines)


def _add_llm_usage(task_id: str, prompt_tokens: int, completion_tokens: int) -> None:
    task = get_task(task_id) or {}
    metrics = task.get("cost_metrics", {})
    metrics["llm_tokens_prompt"] = metrics.get("llm_tokens_prompt", 0) + prompt_tokens
    metrics["llm_tokens_completion"] = metrics.get("llm_tokens_completion", 0) + completion_tokens
    update_task(task_id, {"cost_metrics": metrics})


def _generate_service_overview(task_id: str, service: str, client: str,
                               items: List[Dict[str, Any]]) -> str:
    """One short AI call per service, using a compact resource list (not the full JSON)."""
    fallback = f"{service} is used in {client}'s environment for the resources listed in this section."

    resource_lines = []
    for r in items:
        sku = (r.get("sku") or {}).get("name") or "-"
        resource_lines.append(
            f"- {_display_name(r)} (location: {_loc(r) or '-'}, sku: {sku}, kind: {r.get('kind') or '-'})"
        )
    user_prompt = (
        f"Service: {service}\n"
        f"Client: {client}\n"
        f"Resources ({len(items)}):\n" + "\n".join(resource_lines)
    )

    base_endpoint = AGENT_OPENAI_ENDPOINT.split("/openai")[0].rstrip("/")
    try:
        resp = requests.post(
            f"{base_endpoint}/openai/v1/chat/completions",
            headers={"api-key": AGENT_OPENAI_KEY, "Content-Type": "application/json"},
            json={
                "model": WRITER_DEPLOYMENT,
                "messages": [
                    {"role": "system", "content": OVERVIEW_SYSTEM_PROMPT},
                    {"role": "user", "content": user_prompt},
                ],
                "temperature": 0.2,
                "max_tokens": 300,
            },
            timeout=60,
        )
        if resp.status_code != 200:
            print(f"[overview] {service}: HTTP {resp.status_code}; using fallback")
            return fallback

        data = resp.json()
        usage = data.get("usage", {})
        _add_llm_usage(task_id, usage.get("prompt_tokens", 0), usage.get("completion_tokens", 0))

        text = data["choices"][0]["message"]["content"]
        # Keep plain prose only — drop any heading/list/table lines the model might add
        lines = [l.strip() for l in text.splitlines()
                 if l.strip() and not l.strip().startswith(("#", "|", "-", "*", "```"))]
        overview = " ".join(lines).strip()
        return overview or fallback
    except Exception as e:
        print(f"[overview] {service}: {e}; using fallback")
        return fallback


def build_final_config_markdown(task_id: str, resources: List[Dict[str, Any]]) -> str:
    """Build the document body deterministically from the fetched ARM resources."""
    task = get_task(task_id) or {}
    client = (task.get("clientName") or "the client").strip()
    catalog_keys = list(SERVICE_CATALOG.keys())

    # Group resources by service, sorted by category order then catalog order
    groups: Dict[tuple, List[Dict[str, Any]]] = {}
    for r in resources:
        rtype = (r.get("type") or "").lower()
        if rtype in SERVICE_CATALOG:
            service, category = SERVICE_CATALOG[rtype]
            order = catalog_keys.index(rtype)
        else:
            service, category, order = (r.get("type") or "Other Resource"), "Other", len(catalog_keys)
        key = (CATEGORY_ORDER.index(category), order, service)
        groups.setdefault(key, []).append(r)

    sections: List[str] = []
    table_no = 0

    for key in sorted(groups):
        service = key[2]
        items = sorted(groups[key], key=lambda r: _display_name(r).lower())

        overview = _generate_service_overview(task_id, service, client, items)

        names = [f"**{_display_name(r)}**" for r in items]
        name_list = names[0] if len(names) == 1 else ", ".join(names[:-1]) + " and " + names[-1]
        verb = "resource is" if len(items) == 1 else "resources are"

        block = [
            f"# {service}", "",
            "## Overview", "",
            overview, "",
            f"## {service} Configuration", "",
            f"The following {service} {verb} configured in {client}'s environment: {name_list}.", "",
        ]

        for r in items:
            name = _display_name(r)
            block += [f"### {name}", ""]
            for blade, rows in get_config_blades(r):
                table_no += 1
                block += [
                    f"The table below shows the {blade} of **{name}** in {client}'s environment.", "",
                    _config_table_md(rows), "",
                    f"*Table {table_no} {name} {blade}*", "",
                ]

        sections.append("\n".join(block))
        print(f"[build_markdown] {service}: {len(items)} resource(s)")

    return "\n".join(sections).strip() + "\n"


def background_final_config_processing(task_id: str, token: str, resource_ids: List[str]):
    task = get_task(task_id)
    if not task:
        return
        
    task_dir = UPLOAD_DIR / task_id

    try:
        update_task(task_id, {
            "step_name": "Fetching Azure Resources...",
            "progress": 20
        })

        fetch_result = fetch_resource_details(resource_ids, token)
        resources_data = fetch_result["succeeded"]   # ONLY this goes to the AI
        fetch_errors = fetch_result["failed"]

        # Save both lists for debugging
        raw_json_path = task_dir / "azure_resources.json"
        with open(raw_json_path, "w", encoding="utf-8") as f:
            json.dump(fetch_result, f, indent=2, default=str)

        # Record counts + errors on the task (no token is stored)
        latest_task = get_task(task_id) or {}
        cost_metrics = latest_task.get("cost_metrics", {})
        cost_metrics["resources_fetched"] = len(resources_data)
        cost_metrics["resources_failed"] = len(fetch_errors)
        update_task(task_id, {"cost_metrics": cost_metrics, "fetch_errors": fetch_errors})

        # Stop early with a clear error instead of generating an empty/false document
        if not resources_data:
            first_reason = fetch_errors[0]["reason"] if fetch_errors else "unknown"
            raise RuntimeError(
                f"None of the {len(resource_ids)} selected resources could be fetched. "
                f"First error: {first_reason}"
            )

        update_task(task_id, {
            "step_name": "Drafting Document with AI (Taking 1-2 mins)...",
            "progress": 50
        })

        markdown_content = build_final_config_markdown(task_id, resources_data)
        
        md_file = task_dir / "FINAL_CONFIG_DOCUMENT.md"
        with open(md_file, "w", encoding="utf-8") as f:
            f.write(markdown_content)

        update_task(task_id, {
            "step_name": "Converting to Word Document...",
            "progress": 85
        })

        docx_path = task_dir / "Final_Configuration_Document.docx"
        client_name = task.get("clientName", "").strip()
        project_title = task.get("projectTitle", "Project").strip()
        logo_path = task.get("logo_path")

        build_final_config_docx(md_file, docx_path, project_title, client_name, logo_path)

        blob_res = {"uploaded": False, "files": {}}
        try:
            blob_res = upload_task_to_blob(task_id, task_dir, files=[
                "Final_Configuration_Document.docx",
                "FINAL_CONFIG_DOCUMENT.md",
                "azure_resources.json",
            ])
        except Exception as e:
            print(f"[background_final_config] Blob upload failed (non-critical): {e}")

        docx_blob_name = (blob_res.get("files") or {}).get("docx") or f"{task_id}/Final_Configuration_Document.docx"

        update_task(task_id, {
            "status": "completed",
            "step_name": "Ready!",
            "progress": 100,
            "blobPath": docx_blob_name,
            "storage": blob_res,
            "generated_filename": "generated_" + re.sub(r"[^A-Za-z0-9]+", "_", f"{project_title}_{client_name}").strip("_")
        })

    except Exception as e:
        error_trace = traceback.format_exc()
        print(f"[background_final_config] FAILED for task {task_id}:")
        print(error_trace)
        update_task(task_id, {
            "status": "failed",
            "step_name": "Failed",
            "progress": 0,
            "error": str(e),
            "error_trace": error_trace
        })
