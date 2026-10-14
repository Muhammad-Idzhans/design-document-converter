"""
task_manager.py — Task lifecycle management: CRUD operations, Cosmos DB sync, and Blob Storage uploads.

Handles the hybrid task store (local JSON file + Azure Cosmos DB) and 
Azure Blob Storage uploads for generated documents.
"""

import os
import json
import shutil
import datetime
from typing import Optional, Dict, Any, List

from config import (
    TASKS_FILE, TASKS_LOCK, UPLOAD_DIR,
    cosmos_tasks_container, BlobServiceClient, AZURE_BLOB_CONTAINER_NAME
)
from pricing import compute_cost_summary

# Fields to remove from Cosmos DB documents before upserting (large or ephemeral)
COSMOS_EXCLUDED_FIELDS = {
    "markdown_draft",
    "asset_library",
    "preview_data",
    "file_path",
    "logo_path",
    "result_docx",
}


# -----------------------------
# Task Store (Azure-safe + less corruption risk)
# -----------------------------
def load_tasks() -> Dict[str, Any]:
    with TASKS_LOCK:
        if TASKS_FILE.exists():
            try:
                with open(TASKS_FILE, "r", encoding="utf-8") as f:
                    return json.load(f)
            except Exception as e:
                print(f"[load_tasks] Failed to parse tasks_db.json: {e}")
                return {}
        return {}


def save_tasks(tasks_dict: Dict[str, Any]) -> None:
    with TASKS_LOCK:
        with open(TASKS_FILE, "w", encoding="utf-8") as f:
            json.dump(tasks_dict, f, indent=4, ensure_ascii=False, default=str)


def finalize_task_record(task_id: str, raw_task: Dict[str, Any]) -> Dict[str, Any]:
    """
    Prepare a lean, locked metadata document for Cosmos DB:
    1. Attach locked cost_summary if not already present
    2. Ensure blobPath & storage metadata are defined
    3. Truncate long error messages / arrays
    4. Strip heavy / local fields (markdown_draft, preview_data, etc.)
    """
    task_record = dict(raw_task)
    task_record["id"] = task_id
    task_record["taskId"] = task_id

    # 1. Lock generation cost snapshot if missing
    if "cost_summary" not in task_record or not task_record["cost_summary"]:
        cost_summary = compute_cost_summary(task_record.get("cost_metrics"))
        if cost_summary:
            task_record["cost_summary"] = cost_summary

    # 2. Ensure blobPath is set
    if not task_record.get("blobPath"):
        storage_files = (task_record.get("storage") or {}).get("files", {})
        if "docx" in storage_files:
            task_record["blobPath"] = storage_files["docx"]
        else:
            doc_type = task_record.get("documentType", "")
            if doc_type == "Final Configuration Document":
                task_record["blobPath"] = f"{task_id}/Final_Configuration_Document.docx"
            else:
                task_record["blobPath"] = f"{task_id}/Solution_Design_Document.docx"

    # 3. Truncate lengthy error details
    if "error_trace" in task_record and isinstance(task_record["error_trace"], str):
        task_record["error_trace"] = task_record["error_trace"][:4000]

    if "fetch_errors" in task_record and isinstance(task_record["fetch_errors"], list):
        task_record["fetch_errors"] = task_record["fetch_errors"][:50]

    # 4. Strip excluded fields from Cosmos DB
    for key in COSMOS_EXCLUDED_FIELDS:
        task_record.pop(key, None)

    return task_record


def update_task(task_id: str, updates: Dict[str, Any]) -> None:
    with TASKS_LOCK:
        tasks = load_tasks()
        if task_id not in tasks:
            tasks[task_id] = {}
        tasks[task_id].update(updates)
        save_tasks(tasks)
        
        # Phase 3: Hybrid Cosmos DB approach
        status = tasks[task_id].get("status")
        if status in ["completed", "failed"] and cosmos_tasks_container is not None:
            try:
                task_record = finalize_task_record(task_id, tasks[task_id])
                cosmos_tasks_container.upsert_item(task_record)
                print(f"[CosmosDB] Upserted final lean task {task_id} successfully.")
                
                # Phase 4: Automatic Server Cleanup
                task_dir = UPLOAD_DIR / task_id
                if task_dir.exists():
                    shutil.rmtree(task_dir, ignore_errors=True)
                    print(f"[Cleanup] Deleted local files for task {task_id}")
                    
                if task_id in tasks:
                    del tasks[task_id]
                    save_tasks(tasks)
                    print(f"[Cleanup] Removed task {task_id} from local tasks_db.json")
                    
            except Exception as e:
                print(f"[CosmosDB/Cleanup] Failed to upsert/cleanup task {task_id}: {e}")


def get_task(task_id: str) -> Optional[Dict[str, Any]]:
    tasks = load_tasks()
    task = tasks.get(task_id)
    if task:
        return task
        
    # Phase 3 fallback: if not in local JSON, try Cosmos DB
    if cosmos_tasks_container is not None:
        try:
            response = cosmos_tasks_container.read_item(item=task_id, partition_key=task_id)
            return response
        except Exception:
            # Usually means Not Found (404)
            pass
            
    return None


# -----------------------------
# Azure Blob Storage Helper
# -----------------------------
def upload_task_to_blob(task_id: str, task_dir, files: Optional[List[str]] = None) -> Dict[str, Any]:
    """
    Upload key task outputs to Azure Blob Storage after generation.
    Returns details of uploaded files for recording on the task document.
    """
    conn_str = os.getenv("AZURE_STORAGE_CONNECTION_STRING")
    container = os.getenv("AZURE_BLOB_CONTAINER_NAME", AZURE_BLOB_CONTAINER_NAME)
    
    result = {
        "uploaded": False,
        "container": container,
        "files": {},
        "uploaded_at": datetime.datetime.utcnow().isoformat() + "Z"
    }

    if not conn_str:
        print("[upload_task_to_blob] No connection string, skipping blob upload")
        return result
    
    if BlobServiceClient is None:
        print("[upload_task_to_blob] BlobServiceClient not available, skipping blob upload")
        return result
    
    try:
        blob_service = BlobServiceClient.from_connection_string(conn_str)
        container_client = blob_service.get_container_client(container)
        
        if files is None:
            # Default: Design Document outputs + slide thumbnails
            files_to_upload = ["Solution_Design_Document.docx", "FINAL_DESIGN_DOCUMENT.md", "extraction_payload.json"]
            files_to_upload += [f.name for f in task_dir.glob("slide_*.jpg")]
        else:
            files_to_upload = list(files)
        
        uploaded_map = {}
        for filename in files_to_upload:
            filepath = task_dir / filename
            if filepath.exists():
                blob_name = f"{task_id}/{filename}"
                with open(filepath, "rb") as data:
                    container_client.upload_blob(blob_name, data, overwrite=True)
                
                # Classify key files for clean referencing
                if filename.endswith(".docx"):
                    uploaded_map["docx"] = blob_name
                elif filename.endswith(".md"):
                    uploaded_map["markdown"] = blob_name
                elif filename == "extraction_payload.json" or filename == "azure_resources.json":
                    uploaded_map["source_data"] = blob_name

        result["uploaded"] = True
        result["files"] = uploaded_map
        print(f"[upload_task_to_blob] Uploaded outputs for task {task_id}: {list(uploaded_map.keys())}")
    except Exception as e:
        print(f"[upload_task_to_blob] Upload failed (non-critical): {e}")

    return result
