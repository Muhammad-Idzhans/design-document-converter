"""
app_dev.py — FastAPI application entry point.

This is the slim entry point that defines all HTTP endpoints and middleware.
All business logic has been moved to dedicated modules:
- config.py          → Environment variables, paths, constants
- task_manager.py    → Task CRUD, Cosmos DB sync, Blob storage
- doc_styles.py      → Shared DOCX formatting pipeline
- design_doc_agents.py → Design Document AI agents and processing
- final_config_agents.py → (Placeholder) Final Config Document agents

The original monolithic file is preserved as app_dev_original.py for reference.
"""

import os
import uuid
import json
import shutil
import datetime
from pathlib import Path
from typing import Optional
from io import BytesIO

from fastapi import FastAPI, UploadFile, File, Form, BackgroundTasks, HTTPException, Header
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, StreamingResponse

from config import (
    AGENT_OPENAI_ENDPOINT, AGENT_ASSISTANT_ID,
    UPLOAD_DIR, ENV_PREFIX,
    _get_soffice, safe_filename, _parse_cors_origins,
    ProcessRequest, cosmos_tasks_container, BlobServiceClient,
    AZURE_BLOB_CONTAINER_NAME
)
from task_manager import update_task, get_task
from design_doc_agents import process_initial_upload, background_processing
from final_config_agents import get_azure_subscriptions, get_azure_resource_groups, get_azure_resources, background_final_config_processing


# ──────────────────────────────────────────────────────
# FastAPI App & Middleware
# ──────────────────────────────────────────────────────
app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=_parse_cors_origins(),
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


# ──────────────────────────────────────────────────────
# Startup Diagnostics
# ──────────────────────────────────────────────────────
@app.on_event("startup")
async def test_agent_connection():
    print(f"\n{'='*60}")
    print(f"  DESIGN DOCUMENT CONVERTER - STARTUP DIAGNOSTICS")
    print(f"{'='*60}")

    # Verify LibreOffice
    try:
        soffice = _get_soffice()
        print(f"  [OK] LibreOffice found at: {soffice}")
    except RuntimeError as e:
        print(f"  [WARN] LibreOffice NOT found: {e}")
        print(f"         PDF conversion and TOC update will fail.")

    # Test Agent connection
    if AGENT_ASSISTANT_ID:
        try:
            from azure.identity import DefaultAzureCredential
            from azure.ai.projects import AIProjectClient
            project_client = AIProjectClient(
                endpoint=AGENT_OPENAI_ENDPOINT,
                credential=DefaultAzureCredential()
            )
            client = project_client.get_openai_client()

            agent_parts = AGENT_ASSISTANT_ID.split(":")
            agent_name = agent_parts[0]
            agent_version = agent_parts[1] if len(agent_parts) > 1 else "1"

            test_response = client.responses.create(
                input=[{"type": "message", "role": "user", "content": "Say 'Agent is online' in one sentence."}],
                extra_body={"agent_reference": {"name": agent_name, "version": agent_version, "type": "agent_reference"}}
            )
            print(f"  [OK] Agent Connection Successful!")
            print(f"       Agent Response: {test_response.output_text}")
        except Exception as e:
            print(f"  [WARN] Agent Connection Test Failed: {e}")
            print(f"         The system will fall back to direct LLM calls.")
    else:
        print(f"  [INFO] AGENT_ASSISTANT_ID not set. Agent features disabled, using direct LLM fallback.")

    print(f"{'='*60}\n")


# ──────────────────────────────────────────────────────
# Health Check
# ──────────────────────────────────────────────────────
@app.get("/")
async def health_check():
    return {"status": "ok", "message": "Design Document Converter API is running."}

@app.get("/api/health")
async def health():
    return {"status": "ok"}


# ──────────────────────────────────────────────────────
# Design Document Endpoints
# ──────────────────────────────────────────────────────

@app.post("/api/upload")
async def upload_file(
    background_tasks: BackgroundTasks,
    file: UploadFile = File(...),
    logo: Optional[UploadFile] = File(None)
):
    task_id = f"{ENV_PREFIX}{uuid.uuid4().hex[:12]}"
    task_dir = UPLOAD_DIR / task_id
    task_dir.mkdir(parents=True, exist_ok=True)

    src_name = safe_filename(file.filename or "source.pptx")
    file_path = str(task_dir / f"source_{src_name}")
    with open(file_path, "wb") as f:
        contents = await file.read()
        f.write(contents)

    # Save optional client logo
    logo_path = None
    if logo and logo.filename:
        logo_ext = Path(logo.filename).suffix or ".png"
        logo_dest = task_dir / f"client_logo{logo_ext}"
        with open(logo_dest, "wb") as f:
            logo_bytes = await logo.read()
            f.write(logo_bytes)
        logo_path = str(logo_dest)

    update_task(task_id, {
        "status": "processing_upload",
        "step_name": "Uploading file...",
        "filename": src_name,
        "file_path": file_path,
        "logo_path": logo_path,
        "createdAt": datetime.datetime.utcnow().isoformat(),
        "documentType": "Design Document",
        "cost_metrics": {
            "vision_tokens_prompt": 0,
            "vision_tokens_completion": 0,
            "llm_tokens_prompt": 0,
            "llm_tokens_completion": 0,
            "content_understanding_pages": 0
        }
    })

    background_tasks.add_task(process_initial_upload, task_id, file_path, logo_path, src_name, task_dir)
    return {"task_id": task_id}


@app.get("/api/thumbnails/{task_id}/{filename}")
async def get_thumbnail(task_id: str, filename: str):
    safe_name = safe_filename(filename, "slide.jpg")
    file_path = UPLOAD_DIR / task_id / safe_name
    if not file_path.exists():
        raise HTTPException(status_code=404, detail="Thumbnail not found")
    return FileResponse(str(file_path), media_type="image/jpeg")


@app.post("/api/process/{task_id}")
async def start_processing(
    task_id: str,
    background_tasks: BackgroundTasks,
    body: Optional[ProcessRequest] = None
):
    task = get_task(task_id)
    if not task:
        raise HTTPException(status_code=404, detail="Task not found")

    # Save any user-modified notes
    if body and body.updated_notes:
        task_dir = UPLOAD_DIR / task_id
        notes_path = task_dir / "user_notes.json"
        with open(notes_path, "w", encoding="utf-8") as f:
            json.dump(body.updated_notes, f)

    background_tasks.add_task(background_processing, task_id)
    return {"status": "processing"}


@app.get("/api/status/{task_id}")
async def get_status(task_id: str):
    task = get_task(task_id)
    if not task:
        raise HTTPException(status_code=404, detail="Task not found")
    return task

# Which DOCX to serve, based on the task's documentType (stored in Cosmos DB)
DOCX_FILENAMES = {
    "Final Configuration Document": "Final_Configuration_Document.docx",
}
DEFAULT_DOCX_FILENAME = "Solution_Design_Document.docx"

@app.get("/api/download/{task_id}")
async def download_doc(task_id: str, filename: str = "preview"):
    task = get_task(task_id)
    if not task:
        raise HTTPException(status_code=404, detail="Task not found")

    # 1. Resolve exact blob name from task metadata (or fall back to convention)
    blob_name = task.get("blobPath") or (task.get("storage") or {}).get("files", {}).get("docx")
    if not blob_name:
        docx_name = DOCX_FILENAMES.get(task.get("documentType"), DEFAULT_DOCX_FILENAME)
        blob_name = f"{task_id}/{docx_name}"
    else:
        docx_name = Path(blob_name).name

    # 2. Try local file first (if task was run in current server instance before cleanup)
    local_path = UPLOAD_DIR / task_id / docx_name
    if local_path.exists():
        return FileResponse(
            str(local_path),
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            filename=f"{filename}.docx"
        )

    # 3. Fallback: stream from Azure Blob Storage (supports dev and prod containers)
    conn_str = os.getenv("AZURE_STORAGE_CONNECTION_STRING")
    if conn_str and BlobServiceClient is not None:
        blob_service = BlobServiceClient.from_connection_string(conn_str)
        
        # Build prioritized list of candidate containers
        candidate_containers = []
        task_container = (task.get("storage") or {}).get("container")
        if task_container:
            candidate_containers.append(task_container)

        env_container = os.getenv("AZURE_BLOB_CONTAINER_NAME", AZURE_BLOB_CONTAINER_NAME)
        if env_container not in candidate_containers:
            candidate_containers.append(env_container)

        for fallback in ["prod-enfrasysdocgen-storage", "dev-enfrasysdocgen-storage"]:
            if fallback not in candidate_containers:
                candidate_containers.append(fallback)

        for c_name in candidate_containers:
            try:
                blob_client = blob_service.get_blob_client(c_name, blob_name)
                if blob_client.exists():
                    blob_data = blob_client.download_blob().readall()
                    return StreamingResponse(
                        BytesIO(blob_data),
                        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        headers={"Content-Disposition": f'attachment; filename="{filename}.docx"'}
                    )
            except Exception as e:
                print(f"[download_doc] Blob search error in {c_name} for {blob_name}: {e}")

    raise HTTPException(status_code=404, detail="Document file not found (not local, not in blob)")


@app.delete("/api/tasks/{task_id}")
async def delete_task_endpoint(task_id: str):
    results = {"cosmos_soft_deleted": False, "blob_deleted_count": 0, "local_deleted": False}
    now_iso = datetime.datetime.now(datetime.timezone.utc).isoformat()

    # 1. Soft-delete in Cosmos DB: preserve financial and token analytics, set isDeleted = True
    if cosmos_tasks_container is not None:
        try:
            item = cosmos_tasks_container.read_item(item=task_id, partition_key=task_id)
            item["isDeleted"] = True
            item["deletedAt"] = now_iso
            item["blobPath"] = None
            cosmos_tasks_container.upsert_item(item)
            results["cosmos_soft_deleted"] = True
            print(f"[delete_task] Soft-deleted task {task_id} in Cosmos DB (isDeleted=True).")
        except Exception as e:
            print(f"[delete_task] Cosmos DB soft delete failed for {task_id}: {e}")

    # 2. Delete from Azure Blob Storage (check both dev and prod containers)
    conn_str = os.getenv("AZURE_STORAGE_CONNECTION_STRING")
    if conn_str and BlobServiceClient is not None:
        try:
            blob_service = BlobServiceClient.from_connection_string(conn_str)
            for c_name in ["dev-enfrasysdocgen-storage", "prod-enfrasysdocgen-storage"]:
                try:
                    container_client = blob_service.get_container_client(c_name)
                    blobs = list(container_client.list_blobs(name_starts_with=f"{task_id}/"))
                    for blob in blobs:
                        container_client.delete_blob(blob.name)
                        results["blob_deleted_count"] += 1
                except Exception:
                    pass
            print(f"[delete_task] Deleted {results['blob_deleted_count']} blobs for task {task_id}.")
        except Exception as e:
            print(f"[delete_task] Blob delete failed for {task_id}: {e}")

    # 3. Clean up local files
    task_dir = UPLOAD_DIR / task_id
    if task_dir.exists():
        shutil.rmtree(task_dir, ignore_errors=True)
        results["local_deleted"] = True

    return {"status": "deleted", "task_id": task_id, "details": results}

# ──────────────────────────────────────────────────────
# Azure Resource Proxy Endpoints (Final Config)
# ──────────────────────────────────────────────────────

def _get_token_from_header(authorization: Optional[str] = Header(None)) -> str:
    if not authorization:
        raise HTTPException(status_code=401, detail="Authorization header missing")
    return authorization

def _handle_azure_error(e: Exception):
    print(f"[Azure Proxy] ERROR: {e}")
    if isinstance(e, ValueError) and "Authentication Error" in str(e):
        raise HTTPException(status_code=401, detail="Your Azure session has expired. Please sign out and sign back in.")
    raise HTTPException(status_code=500, detail=str(e))

@app.get("/api/azure/subscriptions")
async def api_list_subscriptions(authorization: Optional[str] = Header(None)):
    token = _get_token_from_header(authorization)
    print(f"[Azure Proxy] Token received (first 20 chars): {token[:20]}...")
    try:
        result = get_azure_subscriptions(token)
        print(f"[Azure Proxy] Successfully fetched {len(result)} subscriptions")
        return result
    except Exception as e:
        _handle_azure_error(e)

@app.get("/api/azure/subscriptions/{subscription_id}/resourcegroups")
async def api_list_resource_groups(subscription_id: str, authorization: Optional[str] = Header(None)):
    token = _get_token_from_header(authorization)
    try:
        return get_azure_resource_groups(token, subscription_id)
    except Exception as e:
        _handle_azure_error(e)

@app.get("/api/azure/subscriptions/{subscription_id}/resourcegroups/{rg_name}/resources")
async def api_list_resources(subscription_id: str, rg_name: str, authorization: Optional[str] = Header(None)):
    token = _get_token_from_header(authorization)
    try:
        return get_azure_resources(token, subscription_id, rg_name)
    except Exception as e:
        _handle_azure_error(e)


@app.post("/api/final-config/generate")
async def generate_final_config(
    background_tasks: BackgroundTasks,
    projectTitle: str = Form(...),
    clientName: str = Form(...),
    tenantId: str = Form(...),
    armToken: str = Form(...),
    checkedKeys: str = Form(...),
    logo: Optional[UploadFile] = File(None)
):
    task_id = f"{ENV_PREFIX}{uuid.uuid4().hex[:12]}"
    task_dir = UPLOAD_DIR / task_id
    task_dir.mkdir(parents=True, exist_ok=True)

    # Save optional client logo
    logo_path = None
    if logo and logo.filename:
        logo_ext = Path(logo.filename).suffix or ".png"
        logo_dest = task_dir / f"client_logo{logo_ext}"
        with open(logo_dest, "wb") as f:
            logo_bytes = await logo.read()
            f.write(logo_bytes)
        logo_path = str(logo_dest)

    try:
        keys_list = json.loads(checkedKeys)
    except Exception:
        keys_list = []

    update_task(task_id, {
        "status": "processing",
        "step_name": "Initializing generation...",
        "createdAt": datetime.datetime.utcnow().isoformat(),
        "documentType": "Final Configuration Document",
        "projectTitle": projectTitle,
        "clientName": clientName,
        "tenantId": tenantId,
        "logo_path": logo_path,
        "filename": "Azure Configurations",
        "cost_metrics": {
            "llm_tokens_prompt": 0,
            "llm_tokens_completion": 0,
            "resources_fetched": 0
        }
    })

    background_tasks.add_task(background_final_config_processing, task_id, armToken, keys_list)
    return {"task_id": task_id}


# ──────────────────────────────────────────────────────
# Entry point
# ──────────────────────────────────────────────────────
if __name__ == "__main__":
    import uvicorn
    uvicorn.run("app_dev:app", host="0.0.0.0", port=8000, reload=True)