"""
Formatly API Server
-------------------
This module implements the FastAPI backend for the Formatly service.
It exposes endpoints for document upload, processing status, and retrieval.

Architecture:
    - Imports core logic from the root `core/` directory by modifying `sys.path`.
    - Uses Supabase for storage and database management.
    - orchestrates AI processing tasks (formatting, structure detection).

Note:
    This API shares the `core/` library with the CLI `app.py`.
"""

from fastapi import FastAPI, HTTPException, Depends, Form
from fastapi.middleware.cors import CORSMiddleware
from fastapi.security import HTTPBearer, HTTPAuthorizationCredentials
from pydantic import BaseModel, Field
from typing import Optional, Dict, Any
import sys
import uuid
import time
import base64
import os
import shutil
import tempfile
import traceback
from pathlib import Path
from datetime import datetime, timezone
import asyncio
import jwt
from supabase import create_client, Client
import logging
from dotenv import load_dotenv

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Load environment variables
load_dotenv()

# Calculate root path: v1 -> api -> project_root
sys.path.append(str(Path(__file__).resolve().parent.parent.parent))

# Import Core Logic
from core.api_clients import HuggingFaceClient, GeminiClient
from core.formatter import AdvancedFormatter
from core.style_guides import STYLE_GUIDES
from utils.track_changes import TrackChanges

# Initialize FastAPI app
app = FastAPI(
    title="Formatly API",
    description="Document formatting service API",
    version="1.0.0"
)

# CORS middleware
app.add_middleware(
    CORSMiddleware,
    allow_origins=["https://formatly-v15.vercel.app"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

SUPABASE_URL = os.getenv("SUPABASE_URL")
SUPABASE_SERVICE_ROLE_KEY = os.getenv("SUPABASE_SERVICE_ROLE_KEY")
SUPABASE_JWT_SECRET = os.getenv("SUPABASE_JWT_SECRET")
SUPABASE_ANON_KEY = os.getenv("SUPABASE_ANON_KEY")

if not all([SUPABASE_URL, SUPABASE_SERVICE_ROLE_KEY, SUPABASE_JWT_SECRET, SUPABASE_ANON_KEY]):
    logger.critical("Missing required Supabase credentials. API will not function correctly.")

# Create Supabase client if URL/KEY are present
if SUPABASE_URL and SUPABASE_SERVICE_ROLE_KEY:
    supabase: Client = create_client(SUPABASE_URL, SUPABASE_SERVICE_ROLE_KEY)
else:
    supabase = None
    logger.warning("Supabase credentials missing. Database operations will fail.")

# Security
security = HTTPBearer(auto_error=False)

# Pydantic models

class CreateUploadResponse(BaseModel):
    success: bool
    job_id: str
    upload_url: str
    upload_token: str
    file_path: str
    message: str
    upload_headers: Dict[str, str]

class WebhookUploadComplete(BaseModel):
    job_id: str
    file_path: str
    success: bool

class DocumentStatusResponse(BaseModel):
    success: bool
    job_id: str
    status: str
    progress: int
    result_url: Optional[str] = None
    error: Optional[str] = None

class FormattedDocumentResponse(BaseModel):
    success: bool
    filename: str
    content: str  # base64 encoded
    tracked_changes_content: Optional[str] = None
    metadata: Dict[str, Any] = Field(default_factory=dict)

class ProcessDocumentRequest(BaseModel):
    filename: str
    content: str  # base64 encoded
    style: str
    englishVariant: str
    trackedChanges: bool = False
    options: Dict[str, Any] = Field(default_factory=dict)

class ProcessDocumentResponse(BaseModel):
    success: bool
    job_id: str
    status: str
    message: str

class FormattingStyle(BaseModel):
    id: str
    name: str
    description: str

class EnglishVariant(BaseModel):
    id: str
    name: str
    description: str

VALID_STYLES = set(STYLE_GUIDES.keys())
MAX_UPLOAD_SIZE_BYTES = 50 * 1024 * 1024  # 50MB

FORMATTING_STYLES = [
    {"id": key, "name": f"{key.upper()} Style", "description": val.get("meta", {}).get("description", f"{key.upper()} formatting style")}
    for key, val in STYLE_GUIDES.items()
]

ENGLISH_VARIANTS = [
    {"id": "us", "name": "US English", "description": "American English"},
    {"id": "uk", "name": "UK English", "description": "British English"},
    {"id": "ca", "name": "Canadian English", "description": "Canadian English"},
    {"id": "au", "name": "Australian English", "description": "Australian English"},
]

# Helper functions
def verify_token(credentials: HTTPAuthorizationCredentials = Depends(security)):
    """Verify JWT token with Supabase"""
    if not credentials:
        raise HTTPException(status_code=401, detail="Authorization header required")
    
    try:
        token = credentials.credentials
        payload = jwt.decode(
            token, 
            SUPABASE_JWT_SECRET, 
            algorithms=["HS256"],
            audience="authenticated"
        )
        
        user_id = payload.get("sub")
        if not user_id:
            raise HTTPException(status_code=401, detail="Invalid token: missing user ID")
            
        return {"user_id": user_id, "email": payload.get("email")}
        
    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"Token verification error: {str(e)}")
        raise HTTPException(status_code=401, detail="Token verification failed")

async def generate_signed_upload_url(filename: str, user_id: str) -> Dict[str, str]:
    """Generate signed upload URL for document upload to Supabase Storage"""
    try:
        file_extension = filename.split('.')[-1].lower() if '.' in filename else 'docx'
        timestamp = int(time.time())
        unique_filename = f"documents/{user_id}/{timestamp}_{uuid.uuid4()}.{file_extension}"
        
        storage_bucket = "documents"
        response = supabase.storage.from_(storage_bucket).create_signed_upload_url(unique_filename)
        
        if isinstance(response, dict):
            signed_url = response.get("signedURL") or response.get("signed_url") or response.get("url")
            token = response.get("token", "")
        else:
            signed_url = str(response) if response else None
            token = ""
            
        if not signed_url:
             raise Exception("Failed to generate signed upload URL")

        upload_headers = {
            "Authorization": f"Bearer {SUPABASE_ANON_KEY}",
            "x-upsert": "true"
        }
        
        return {
            "upload_url": signed_url,
            "file_path": unique_filename,
            "upload_token": token,
            "upload_headers": upload_headers
        }
        
    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"Error generating signed upload URL: {str(e)}")
        raise HTTPException(status_code=500, detail="Failed to generate upload URL")

async def create_document_record(user_id: str, filename: str, file_path: str, job_id: str, style: str, language_variant: str, options: Dict[str, Any], file_size: Optional[int] = None) -> str:
    """Create document record in Supabase database"""
    try:
        document_data = {
            "id": job_id,
            "user_id": user_id,
            "filename": filename,
            "original_filename": filename,
            "status": "draft",
            "style_applied": style,
            "language_variant": language_variant,
            "storage_location": file_path,
            "formatting_options": options,
            "tracked_changes": options.get("trackedChanges", False),
            "created_at": datetime.now(timezone.utc).isoformat(),
            "updated_at": datetime.now(timezone.utc).isoformat()
        }
        
        if file_size is not None:
             document_data["file_size"] = file_size
        
        supabase.table("documents").insert(document_data).execute()
        return job_id
        
    except Exception as e:
        logger.error(f"Error creating document record: {str(e)}")
        raise HTTPException(status_code=500, detail=f"Failed to create document record: {str(e)}")

async def update_document_status(job_id: str, status: str, progress: int = None, processing_log: Dict[str, Any] = None, result_url: str = None, tracked_changes_url: str = None, formatting_time: float = None):
    """Update document status in database"""
    try:
        update_data = {
            "status": status,
            "updated_at": datetime.now(timezone.utc).isoformat()
        }

        if progress is not None:
            if not processing_log:
                processing_log = {}
            processing_log["progress"] = progress
            update_data["processing_log"] = processing_log

        if result_url:
            update_data["result_url"] = result_url

        if tracked_changes_url:
            update_data["tracked_changes_url"] = tracked_changes_url
            update_data["tracked_changes"] = True

        if status == "formatted":
            update_data["processed_at"] = datetime.now(timezone.utc).isoformat()
            if formatting_time is not None:
                update_data["formatting_time"] = formatting_time

        supabase.table("documents").update(update_data).eq("id", job_id).execute()

    except Exception as e:
        logger.error(f"Error updating document status: {str(e)}")

async def track_document_usage(user_id: str):
    """Track document usage in subscriptions table"""
    try:
        # Call the database function to increment document usage
        result = supabase.rpc("increment_document_usage", {"p_user_id": user_id}).execute()
        logger.info(f"Document usage incremented for user: {user_id}")
        return result
    except Exception as e:
        logger.error(f"Error tracking document usage: {str(e)}")
        # Don't raise here to allow flow to continue if tracking fails
        # raise

async def track_storage_usage(user_id: str, storage_mb: float):
    """Track storage usage in subscriptions table"""
    try:
        # Convert MB to GB for storage tracking
        storage_gb = storage_mb / 1024
        
        # Call the database function to update storage usage
        result = supabase.rpc("update_storage_usage", {
            "p_user_id": user_id, 
            "p_storage_gb": storage_gb
        }).execute()
        logger.info(f"Storage usage updated for user: {user_id}, added: {storage_gb}GB")
        return result
    except Exception as e:
        logger.error(f"Error tracking storage usage: {str(e)}")
        # Don't raise here
        # raise

async def process_document_task(job_id: str, file_path: str, style: str, english_variant: str, options: Dict[str, Any], user_id: str):
    """Background task to process the document using Core Formatter"""
    temp_dir = tempfile.mkdtemp()
    local_input_path = Path(temp_dir) / "input.docx"
    local_output_path = Path(temp_dir) / "output.docx"
    
    try:
        # 1. Download file from Supabase
        logger.info(f"Downloading file {file_path} for job {job_id}")
        
        file_content = supabase.storage.from_("documents").download(file_path)
        with open(local_input_path, "wb") as f:
            f.write(file_content)
            
        # 2. Update Status to Processing
        await update_document_status(job_id, "processing", 10, {"message": "File downloaded, initializing AI..."})

        # 3. Initialize AI Client & Formatter
        backend = os.getenv("DEFAULT_BACKEND", "huggingface")
        if backend == "gemini":
            client = GeminiClient()
        else:
            client = HuggingFaceClient()
            
        formatter = AdvancedFormatter(style, ai_client=client, english_variant=english_variant)
        
        # 4. Run Formatting
        await update_document_status(job_id, "processing", 30, {"message": f"Analyzing document structure with {backend}..."})
        
        loop = asyncio.get_running_loop()
        start_time = time.time()
        
        await loop.run_in_executor(None, formatter.format_document, str(local_input_path), str(local_output_path))
        
        duration = time.time() - start_time
        
        # 5. Upload Result to Supabase
        await update_document_status(job_id, "processing", 90, {"message": "Formatting complete. Uploading result..."})
        
        result_filename = f"{user_id}/formatted/{job_id}.docx"
        with open(local_output_path, "rb") as f:
            supabase.storage.from_("documents").upload(
                result_filename,
                f,
                {"content-type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document"}
            )

        # 6. Track Changes if requested
        tracked_changes_url = None
        if options.get("trackedChanges"):
            await update_document_status(job_id, "processing", 95, {"message": "Generating tracked changes..."})
            tracker = TrackChanges(str(local_input_path), str(local_output_path))
            tracked_path = await loop.run_in_executor(None, tracker.compare_docs)

            if tracked_path and os.path.exists(tracked_path):
                tracked_filename = f"{user_id}/tracked/{job_id}_tracked.docx"
                with open(tracked_path, "rb") as f:
                    supabase.storage.from_("documents").upload(
                        tracked_filename,
                        f,
                        {"content-type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document"}
                    )
                tracked_changes_url = tracked_filename
                logger.info(f"Tracked changes uploaded: {tracked_filename}")

        # 7. Finalize Status
        processing_log = {
            "progress": 100,
            "processing_time": duration,
            "backend": backend,
            "message": "Formatting successful"
        }

        await update_document_status(
            job_id, "formatted", 100, processing_log,
            result_url=result_filename,
            tracked_changes_url=tracked_changes_url,
            formatting_time=duration
        )
        
        logger.info(f"Job {job_id} completed successfully.")

    except Exception as e:
        logger.error(f"Job {job_id} failed: {e}")
        traceback.print_exc()
        
        error_msg = str(e)
        if "JSON Parsing Error" in error_msg:
            error_msg = "Formatting failed: AI response was malformed."
        elif "connection" in error_msg.lower():
            error_msg = "Formatting failed: Connection error."
            
        await update_document_status(job_id, "failed", 0, {"error": error_msg})
        
    finally:
        shutil.rmtree(temp_dir)

# API Routes

@app.get("/")
async def root():
    return {"message": "Formatly API", "status": "running"}

@app.post("/api/documents/create-upload")
async def create_upload_url(
    filename: str = Form(...),
    style: str = Form(...),
    englishVariant: str = Form(...),
    trackedChanges: bool = Form(False),
    file_size: Optional[int] = Form(None),
    user: dict = Depends(verify_token)
) -> CreateUploadResponse:
    try:
        if style not in VALID_STYLES:
            raise HTTPException(status_code=400, detail=f"Invalid style. Must be one of: {', '.join(VALID_STYLES)}")
        
        if file_size and file_size > MAX_UPLOAD_SIZE_BYTES:
            raise HTTPException(status_code=400, detail="File size exceeds the 50MB limit.")
        
        job_id = str(uuid.uuid4())
        upload_info = await generate_signed_upload_url(filename, user["user_id"])
        
        options = {"trackedChanges": trackedChanges}
        
        await create_document_record(
            user["user_id"], 
            filename, 
            upload_info["file_path"], 
            job_id, 
            style, 
            englishVariant, 
            options,
            file_size
        )
        
        return CreateUploadResponse(
            success=True,
            job_id=job_id,
            upload_url=upload_info["upload_url"],
            upload_token=upload_info["upload_token"],
            file_path=upload_info["file_path"],
            upload_headers=upload_info["upload_headers"],
            message="Upload URL created."
        )
    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"Create upload error: {e}")
        raise HTTPException(status_code=500, detail="An internal error occurred.")

@app.post("/api/documents/upload-complete")
async def upload_complete_webhook(
    webhook_data: WebhookUploadComplete,
    user: dict = Depends(verify_token)
):
    try:
        job_id = webhook_data.job_id
        
        # Verify job ownership
        response = supabase.table("documents").select("*").eq("id", job_id).eq("user_id", user["user_id"]).execute()
        if not response.data:
            raise HTTPException(status_code=404, detail="Job not found")
            
        document = response.data[0]
        
        if webhook_data.success:
            await update_document_status(job_id, "processing", 0)
            
            # Start Background Processing
            asyncio.create_task(process_document_task(
                job_id, 
                webhook_data.file_path,
                document["style_applied"],
                document["language_variant"],
                document["formatting_options"] or {},
                user["user_id"]
            ))
            
            return {"success": True, "message": "Processing started", "job_id": job_id}
        else:
            await update_document_status(job_id, "failed", 0, {"error": "Upload failed"})
            return {"success": False, "message": "Upload failed recorded"}
            
    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"Webhook error: {e}")
        raise HTTPException(status_code=500, detail="An internal error occurred.")

@app.get("/api/documents/status/{job_id}")
async def get_document_status(job_id: str, user: dict = Depends(verify_token)):
    try:
        response = supabase.table("documents").select("*").eq("id", job_id).eq("user_id", user["user_id"]).execute()
        if not response.data:
            raise HTTPException(status_code=404, detail="Job not found")
            
        doc = response.data[0]
        processing_log = doc.get("processing_log") or {}
        
        return DocumentStatusResponse(
            success=True,
            job_id=job_id,
            status=doc["status"],
            progress=processing_log.get("progress", 0),
            error=processing_log.get("error")
        )
    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"Status error: {e}")
        raise HTTPException(status_code=500, detail="An internal error occurred.")

@app.get("/api/documents/download/{job_id}")
async def download_document(job_id: str, user: dict = Depends(verify_token)):
    try:
        response = supabase.table("documents").select("*").eq("id", job_id).eq("user_id", user["user_id"]).execute()
        if not response.data:
            raise HTTPException(status_code=404, detail="Job not found")
            
        doc = response.data[0]
        if doc["status"] != "formatted":
            raise HTTPException(status_code=400, detail="Document not ready")
            
        result_url = doc.get("result_url")
        if not result_url:
            raise HTTPException(status_code=404, detail="Result file not found")
            
        # Download from storage
        file_content = supabase.storage.from_("documents").download(result_url)
        
        # Base64 encode
        encoded_content = base64.b64encode(file_content).decode()
        
        # Download tracked changes if available
        tracked_content = None
        tracked_url = doc.get("tracked_changes_url")
        if tracked_url:
            try:
                tracked_file_content = supabase.storage.from_("documents").download(tracked_url)
                tracked_content = base64.b64encode(tracked_file_content).decode()
            except Exception as e:
                logger.error(f"Failed to download tracked changes: {e}")

        return FormattedDocumentResponse(
            success=True,
            filename=f"formatted_{doc['filename']}",
            content=encoded_content,
            tracked_changes_content=tracked_content,
            metadata={"style": doc["style_applied"]}
        )
        
    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"Download error: {e}")
        raise HTTPException(status_code=500, detail="An internal error occurred.")

@app.post("/api/documents/upload")
async def upload_document(
    filename: str = Form(...),
    style: str = Form(...),
    englishVariant: str = Form(...),
    trackedChanges: bool = Form(False),
    file_size: Optional[int] = Form(None),
    user: dict = Depends(verify_token)
):
    """EXISTING ENDPOINT: Generate presigned upload URL and create document record - KEPT FOR COMPATIBILITY"""
    try:
        if style not in VALID_STYLES:
            raise HTTPException(status_code=400, detail=f"Invalid style. Must be one of: {', '.join(VALID_STYLES)}")
        
        if file_size and file_size > MAX_UPLOAD_SIZE_BYTES:
            raise HTTPException(status_code=400, detail="File size exceeds the 50MB limit.")
        
        job_id = str(uuid.uuid4())
        
        # Reuse generating logic
        upload_info = await generate_signed_upload_url(filename, user["user_id"])
        
        options = {"trackedChanges": trackedChanges}
        
        # Create document record
        await create_document_record(
            user["user_id"], 
            filename, 
            upload_info["file_path"], 
            job_id, 
            style, 
            englishVariant, 
            options,
            file_size
        )
        
        return {
            "success": True,
            "job_id": job_id,
            "upload_url": upload_info["upload_url"],
            "upload_token": upload_info.get("upload_token", ""),
            "file_path": upload_info["file_path"],
            "upload_headers": upload_info["upload_headers"],
            "message": f"Document queued for {style} formatting. Status: DRAFT. Upload your file to begin processing."
        }
    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"Upload endpoint error: {str(e)}")
        raise HTTPException(status_code=500, detail="An internal error occurred.")

@app.post("/api/documents/process")
async def process_document(
    request: ProcessDocumentRequest,
    user: dict = Depends(verify_token)
):
    """Process a document with formatting (legacy endpoint for backward compatibility)"""
    try:
        job_id = str(uuid.uuid4())
        
        file_path = f"legacy/{user['user_id']}/{job_id}.docx"
        
        uploaded = False
        if request.content:
            try:
                content_bytes = base64.b64decode(request.content)
                supabase.storage.from_("documents").upload(
                    file_path,
                    content_bytes,
                    {"content-type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document"}
                )
                uploaded = True
            except Exception as e:
                logger.error(f"Failed to upload legacy content: {e}")
        
        options = request.options.copy()
        options["trackedChanges"] = request.trackedChanges
        
        await create_document_record(
            user["user_id"],
            request.filename,
            file_path,
            job_id,
            request.style,
            request.englishVariant,
            options
        )
        
        # Trigger processing only if content was successfully uploaded
        if uploaded:
             asyncio.create_task(process_document_task(
                job_id, 
                file_path,
                request.style,
                request.englishVariant,
                options,
                user["user_id"]
            ))

        return ProcessDocumentResponse(
            success=True,
            job_id=job_id,
            status="processing" if uploaded else "draft",
            message=f"Document queued for {request.style} formatting"
        )
    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"Process endpoint error: {str(e)}")
        raise HTTPException(status_code=500, detail="An internal error occurred.")

@app.get("/api/formatting/styles")
async def get_formatting_styles():
    return FORMATTING_STYLES

@app.get("/api/formatting/variants")
async def get_english_variants():
    return ENGLISH_VARIANTS

@app.get("/api/jobs")
async def list_jobs(user: dict = Depends(verify_token)):
    """List jobs for the authenticated user"""
    try:
        response = supabase.table("documents").select("*").eq("user_id", user["user_id"]).order("created_at", desc=True).limit(50).execute()
        return {
            "jobs": response.data,
            "total": len(response.data)
        }
    except Exception as e:
        logger.error(f"List jobs error: {e}")
        return {"jobs": [], "total": 0}

@app.delete("/api/jobs/{job_id}")
async def delete_job(job_id: str, user: dict = Depends(verify_token)):
    try:
        # Check ownership
        response = supabase.table("documents").select("*").eq("id", job_id).eq("user_id", user["user_id"]).execute()
        if not response.data:
            raise HTTPException(status_code=404, detail="Job not found")
            
        doc = response.data[0]
        
        tracked_url = doc.get("tracked_changes_url")
        
        # Delete from storage first
        paths_to_remove = [p for p in [
            doc.get("storage_location"),
            doc.get("result_url"),
            tracked_url
        ] if p]
        if paths_to_remove:
            supabase.storage.from_("documents").remove(paths_to_remove)
            
        # Delete record
        supabase.table("documents").delete().eq("id", job_id).execute()
        
        return {"success": True, "message": "Job deleted"}
    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"Delete job error: {e}")
        raise HTTPException(status_code=500, detail="An internal error occurred.")

@app.get("/api/files")
async def list_files(user: dict = Depends(verify_token)):
    """List files (same as jobs for this architecture)"""
    return await list_jobs(user)

if __name__ == "__main__":
    import uvicorn
    port = int(os.getenv("PORT", 50000))
    uvicorn.run(app, host="0.0.0.0", port=port)
