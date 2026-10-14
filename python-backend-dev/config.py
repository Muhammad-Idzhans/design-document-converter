"""
config.py — Shared configuration, environment variables, paths, and utility functions.

All modules import from this file to get environment settings, path constants,
database clients, and shared utility functions. This is the single source of truth
for all configuration across the application.
"""

import os
import re
import shutil
import tempfile
from pathlib import Path
from threading import RLock
from typing import Optional, Dict, Any, List

from dotenv import load_dotenv
from pydantic import BaseModel

# ✅ LOAD .env FIRST — before ANY os.getenv() calls
load_dotenv()


# -----------------------------
# Pydantic Models (Shared)
# -----------------------------
class ProcessRequest(BaseModel):
    updated_notes: Optional[Dict[str, str]] = None


# -----------------------------
# Cosmos DB Initialization
# -----------------------------
cosmos_client = None
cosmos_tasks_container = None
cosmos_settings_container = None
try:
    from azure.cosmos import CosmosClient
    cosmos_uri = os.getenv("COSMOS_DB_URI")
    cosmos_key = os.getenv("COSMOS_DB_KEY")
    if cosmos_uri and cosmos_key:
        cosmos_client = CosmosClient(cosmos_uri, credential=cosmos_key)
        database = cosmos_client.get_database_client(os.getenv("COSMOS_DB_DATABASE", "design-doc-generator"))
        cosmos_tasks_container = database.get_container_client(os.getenv("COSMOS_DB_TASKS_CONTAINER", "tasks"))
        cosmos_settings_container = database.get_container_client(os.getenv("COSMOS_DB_SETTINGS_CONTAINER", "settings"))
        print("[CosmosDB] Connected to Tasks and Settings Containers.")
    else:
        print("[CosmosDB] URI or Key missing from .env file, skipping Cosmos DB init.")
except ImportError:
    print("[ImportWarning] azure-cosmos package not installed, skipping Cosmos DB init.")
except Exception as e:
    print(f"[CosmosDB] Initialization failed: {e}")


# -----------------------------
# Azure Blob Storage (Optional)
# -----------------------------
AZURE_BLOB_CONTAINER_NAME = os.getenv("AZURE_BLOB_CONTAINER_NAME", "dev-enfrasysdocgen-storage")
try:
    from azure.storage.blob import BlobServiceClient
except ImportError:
    BlobServiceClient = None
    print("[ImportWarning] azure.storage.blob package not installed. Blob upload will be unavailable.")


# -----------------------------
# OpenAI SDK (Optional — kept for dependency verification)
# -----------------------------
try:
    from openai import AzureOpenAI  # noqa: F401
except ImportError:
    print("[ImportWarning] AzureOpenAI package not installed, some features may be unavailable.")


# -----------------------------
# Environment & Azure AI Configuration
# -----------------------------
CONTENT_UNDERSTANDING_ENDPOINT = os.getenv("CONTENT_UNDERSTANDING_ENDPOINT", "").rstrip("/")
CONTENT_UNDERSTANDING_KEY = os.getenv("CONTENT_UNDERSTANDING_KEY", "")
API_VERSION = "2025-11-01"

AZURE_OPENAI_ENDPOINT = os.getenv("AZURE_OPENAI_ENDPOINT", "").rstrip("/")
AZURE_OPENAI_KEY = os.getenv("AZURE_OPENAI_KEY", "")
AZURE_OPENAI_DEPLOYMENT_NAME = os.getenv("AZURE_OPENAI_DEPLOYMENT_NAME", "gpt-4o")

AGENT_OPENAI_ENDPOINT = os.getenv("AGENT_OPENAI_ENDPOINT", "").rstrip("/")
AGENT_OPENAI_KEY = os.getenv("AGENT_OPENAI_KEY", "")

ORCHESTRATOR_DEPLOYMENT = os.getenv("ORCHESTRATOR_DEPLOYMENT", "gpt-4.1")
WRITER_DEPLOYMENT = os.getenv("WRITER_DEPLOYMENT", "gpt-4.1")
AGENT_ASSISTANT_ID = os.getenv("AGENT_ASSISTANT_ID", "")
FINAL_CONFIG_AGENT_ID = os.getenv("FINAL_CONFIG_AGENT_ID", "enfrasys-final-config-writer:2")

# Environment prefix for task IDs — isolates dev records from production
# Set ENV_PREFIX="dev-" locally; leave empty in the deployed Web App.
ENV_PREFIX = os.getenv("ENV_PREFIX", "")


# -----------------------------
# Paths & Storage
# -----------------------------
SCRIPT_DIR = Path(__file__).parent

# OUTPUT_DIR for Azure persistent storage
# - In Azure Web App Linux, set OUTPUT_DIR=/home/site/wwwroot/outputs and enable WEBSITES_ENABLE_APP_SERVICE_STORAGE=true
DEFAULT_OUTPUT_DIR = os.getenv("OUTPUT_DIR", str(SCRIPT_DIR / "outputs"))
BASE_STORAGE = Path(DEFAULT_OUTPUT_DIR)
BASE_STORAGE.mkdir(parents=True, exist_ok=True)

UPLOAD_DIR = BASE_STORAGE / "api_uploads"
UPLOAD_DIR.mkdir(parents=True, exist_ok=True)

TASKS_FILE = BASE_STORAGE / "tasks_db.json"
TASKS_LOCK = RLock()


# -----------------------------
# LibreOffice Path Detection
# -----------------------------
def _find_soffice() -> str:
    """Locate LibreOffice soffice binary. Works on both Linux and Windows."""
    # 1. Check if soffice is on PATH (Linux containers, or Windows with PATH configured)
    found = shutil.which("soffice")
    if found:
        return found

    # 2. Windows: check common installation paths
    if os.name == "nt":
        candidates = [
            r"C:\Program Files\LibreOffice\program\soffice.exe",
            r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        ]
        for candidate in candidates:
            if os.path.isfile(candidate):
                return candidate

    raise RuntimeError(
        "LibreOffice 'soffice' not found. "
        "Install LibreOffice and ensure it is on your PATH or in a standard location."
    )


# Cache the soffice path at module level (evaluated lazily on first call)
_SOFFICE_PATH: Optional[str] = None


def _get_soffice() -> str:
    """Return cached soffice path, finding it on first call."""
    global _SOFFICE_PATH
    if _SOFFICE_PATH is None:
        _SOFFICE_PATH = _find_soffice()
    return _SOFFICE_PATH


def _make_user_profile() -> str:
    """Create a temporary LibreOffice user profile directory.
    
    Each LibreOffice invocation needs its own profile to avoid
    lock conflicts when multiple requests run concurrently.
    """
    return tempfile.mkdtemp(prefix="lo_profile_")


# -----------------------------
# Filename sanitization (security hardening)
# -----------------------------
_FILENAME_SAFE_RE = re.compile(r"[^A-Za-z0-9._-]+")


def safe_filename(name: str, default: str = "file") -> str:
    # Keep base name only (prevents ../../ path traversal)
    base = Path(name).name.strip()
    if not base:
        base = default
    base = _FILENAME_SAFE_RE.sub("_", base)
    # Avoid extremely long names
    return base[:180]


# -----------------------------
# CORS Configuration
# -----------------------------
base_origins = [
    "http://localhost:3000",
    "https://design-document-converter.vercel.app",
    "https://design-document-converter.azurewebsites.net",
    "https://enfrasys-document-generator-bzh5fseuc7b4ehah.southeastasia-01.azurewebsites.net"
]


def _parse_cors_origins() -> List[str]:
    # Start with our known safe origins
    origins = list(base_origins)
    # Allow Azure to configure additional origins via env: CORS_ORIGINS="https://yourfrontend,https://another"
    env_val = os.getenv("CORS_ORIGINS", "").strip()
    if env_val:
        origins.extend([o.strip() for o in env_val.split(",") if o.strip()])
    return origins
