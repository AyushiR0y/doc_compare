from __future__ import annotations
import json
import uuid
from datetime import datetime, timezone
from io import BytesIO
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile
from fastapi import FastAPI, File, HTTPException, Request, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse, StreamingResponse
from comparison_core import ALLOWED_EXTENSIONS, compare_documents, compare_documents_with_preview
from usage_storage import get_usage_log_path
from fastapi.staticfiles import StaticFiles
from fastapi.responses import FileResponse

app = FastAPI(
    title="Document Comparison API",
    description="Compare PDF and DOCX files and download highlighted outputs.",
    version="1.0.0",
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

USAGE_LOG_FILE = get_usage_log_path(__file__)

# Serve frontend static files when the built frontend exists (used in production deployments)
frontend_dist = Path(__file__).parent / "frontend" / "dist"
if frontend_dist.exists():
    app.mount("/static", StaticFiles(directory=str(frontend_dist)), name="static")

    @app.get("/", include_in_schema=False)
    async def root_index():
        index_path = frontend_dist / "index.html"
        if index_path.exists():
            return FileResponse(index_path)
        return JSONResponse({"status": "ok", "service": "document-comparison-api"})


def _extract_client_ip(request: Request) -> str:
    for key in ["x-forwarded-for", "x-real-ip", "cf-connecting-ip", "x-client-ip"]:
        value = request.headers.get(key)
        if value:
            return value.split(",")[0].strip()

    if request.client and request.client.host:
        return request.client.host

    return "unknown"


def _append_usage_event(request: Request, doc1_name: str, doc2_name: str, ext1: str, ext2: str) -> None:
    event = {
        "event_id": str(uuid.uuid4()),
        "event_type": "comparison",
        "timestamp_utc": datetime.now(timezone.utc).isoformat(),
        "comparison_mode": "fastapi",
        "doc1_name": doc1_name,
        "doc2_name": doc2_name,
        "doc1_type": ext1,
        "doc2_type": ext2,
        "upload_count": 2,
        "client_ip": _extract_client_ip(request),
    }

    try:
        USAGE_LOG_FILE.parent.mkdir(parents=True, exist_ok=True)
        with open(USAGE_LOG_FILE, "a", encoding="utf-8") as file:
            file.write(json.dumps(event, ensure_ascii=False) + "\n")
    except OSError:
        pass


def _get_extension(filename: str) -> str:
    return Path(filename).suffix.lower().lstrip(".")


@app.get("/health")
async def health() -> JSONResponse:
    return JSONResponse({"status": "ok", "service": "document-comparison-api"})


@app.post("/compare")
async def compare(request: Request, file1: UploadFile = File(...), file2: UploadFile = File(...)) -> StreamingResponse:
    if not file1.filename or not file2.filename:
        raise HTTPException(status_code=400, detail="Both files must include filenames")

    ext1 = _get_extension(file1.filename)
    ext2 = _get_extension(file2.filename)

    if ext1 not in ALLOWED_EXTENSIONS or ext2 not in ALLOWED_EXTENSIONS:
        raise HTTPException(status_code=400, detail="Only .pdf and .docx files are supported")

    file1_bytes = await file1.read()
    file2_bytes = await file2.read()

    if not file1_bytes or not file2_bytes:
        raise HTTPException(status_code=400, detail="Uploaded files cannot be empty")

    try:
        result = compare_documents(file1.filename, file1_bytes, file2.filename, file2_bytes)
    except ValueError as ex:
        raise HTTPException(status_code=400, detail=str(ex)) from ex
    except Exception as ex:
        raise HTTPException(status_code=500, detail=f"Comparison failed: {str(ex)}") from ex

    _append_usage_event(request, file1.filename, file2.filename, ext1, ext2)

    zip_buffer = BytesIO()
    with ZipFile(zip_buffer, mode="w", compression=ZIP_DEFLATED) as archive:
        archive.writestr(result["doc1_output_name"], result["doc1_bytes"])
        archive.writestr(result["doc2_output_name"], result["doc2_bytes"])
        archive.writestr("summary.json", json.dumps(result["summary"], indent=2))

    zip_buffer.seek(0)

    return StreamingResponse(
        zip_buffer,
        media_type="application/zip",
        headers={"Content-Disposition": "attachment; filename=comparison_result.zip"},
    )


@app.post("/compare-preview")
async def compare_preview(
    request: Request,
    file1: UploadFile = File(...),
    file2: UploadFile = File(...),
    max_pages: int = 0,
    include_images: bool = False,
) -> JSONResponse:
    if not file1.filename or not file2.filename:
        raise HTTPException(status_code=400, detail="Both files must include filenames")

    ext1 = _get_extension(file1.filename)
    ext2 = _get_extension(file2.filename)

    if ext1 not in ALLOWED_EXTENSIONS or ext2 not in ALLOWED_EXTENSIONS:
        raise HTTPException(status_code=400, detail="Only .pdf and .docx files are supported")

    if max_pages < 0:
        raise HTTPException(status_code=400, detail="max_pages must be 0 or greater")

    file1_bytes = await file1.read()
    file2_bytes = await file2.read()

    if not file1_bytes or not file2_bytes:
        raise HTTPException(status_code=400, detail="Uploaded files cannot be empty")

    try:
        preview_result = compare_documents_with_preview(
            file1.filename,
            file1_bytes,
            file2.filename,
            file2_bytes,
            max_pages=max_pages,
            include_images=include_images,
        )
    except ValueError as ex:
        raise HTTPException(status_code=400, detail=str(ex)) from ex
    except Exception as ex:
        raise HTTPException(status_code=500, detail=f"Preview comparison failed: {str(ex)}") from ex

    _append_usage_event(request, file1.filename, file2.filename, ext1, ext2)
    return JSONResponse(preview_result)
