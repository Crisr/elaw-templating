import os
import uuid
import json
import sqlite3
import threading
import time
import shutil
from pathlib import Path
from contextlib import asynccontextmanager

from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import FileResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from typing import Optional

import translate_docx
from worker import TranslationWorker

UPLOAD_DIR = Path("uploads")
DB_PATH = "jobs.db"
MAX_FILE_PAIRS = 10

UPLOAD_DIR.mkdir(exist_ok=True)

# --- SQLite Job Manager ---

_init_lock = threading.Lock()

def _get_db():
    db = sqlite3.connect(DB_PATH, check_same_thread=False)
    db.row_factory = sqlite3.Row
    return db

def _init_db():
    with _init_lock:
        db = _get_db()
        db.execute("""
            CREATE TABLE IF NOT EXISTS jobs (
                id TEXT PRIMARY KEY,
                status TEXT DEFAULT 'pending',
                progress INTEGER DEFAULT 0,
                total INTEGER DEFAULT 0,
                source_file TEXT,
                result_file TEXT,
                error TEXT,
                language TEXT,
                mode TEXT,
                provider TEXT,
                model TEXT,
                created_at TEXT,
                updated_at TEXT
            )
        """)
        db.commit()
        db.close()

def create_job(lang, mode, provider, model):
    job_id = str(uuid.uuid4())
    now = time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime())
    db = _get_db()
    db.execute(
        "INSERT INTO jobs (id, status, language, mode, provider, model, created_at, updated_at) VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
        (job_id, "pending", lang, mode, provider, model, now, now),
    )
    db.commit()
    db.close()
    return job_id

def update_job(job_id, **kwargs):
    sets = ", ".join(f"{k} = ?" for k in kwargs)
    vals = list(kwargs.values()) + [job_id]
    now = time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime())
    db = _get_db()
    db.execute(f"UPDATE jobs SET {sets}, updated_at = ? WHERE id = ?", vals + [now])
    db.commit()
    db.close()

def get_job(job_id):
    db = _get_db()
    row = db.execute("SELECT * FROM jobs WHERE id = ?", (job_id,)).fetchone()
    db.close()
    if row is None:
        return None
    return dict(row)

def enforce_file_limit():
    """Keep at most MAX_FILE_PAIRS pairs (source+result) in uploads dir.
    Delete oldest by mtime when over limit."""
    files = sorted(UPLOAD_DIR.iterdir(), key=lambda f: f.stat().st_mtime)
    pairs = 0
    for f in files:
        if f.name.endswith("_source.docx") or f.name.endswith("_result.docx"):
            pairs += 1
    while pairs > MAX_FILE_PAIRS * 2 and files:
        f = files.pop(0)
        if f.name.endswith("_source.docx") or f.name.endswith("_result.docx"):
            f.unlink()
            pairs -= 1

# --- Config ---

config = translate_docx.load_config()

# --- Worker ---

worker = None

@asynccontextmanager
async def lifespan(app: FastAPI):
    global worker
    _init_db()
    worker = TranslationWorker()
    worker.start()
    yield
    worker.stop()

app = FastAPI(lifespan=lifespan)

# --- API Routes ---

@app.post("/api/translate", status_code=201)
async def api_translate(
    file: UploadFile = File(...),
    lang: str = Form(...),
    mode: str = Form("inline"),
    provider: Optional[str] = Form(None),
    model: Optional[str] = Form(None),
):
    if lang not in ("ro", "en"):
        raise HTTPException(400, "lang must be 'ro' or 'en'")
    if mode not in ("inline", "side-by-side"):
        raise HTTPException(400, "mode must be 'inline' or 'side-by-side'")

    job_id = create_job(lang, mode, provider, model)
    source_path = UPLOAD_DIR / f"{job_id}_source.docx"
    content = await file.read()
    with open(source_path, "wb") as f:
        f.write(content)

    update_job(job_id, source_file=str(source_path))
    worker.enqueue(job_id)

    return {"job_id": job_id, "status": "pending"}

@app.get("/api/translate/{job_id}/status")
def api_status(job_id: str):
    job = get_job(job_id)
    if job is None:
        raise HTTPException(404, "Job not found")
    return {
        "job_id": job["id"],
        "status": job["status"],
        "progress": job["progress"],
        "total": job["total"],
        "error": job["error"],
    }

@app.get("/api/translate/{job_id}/download")
def api_download(job_id: str):
    job = get_job(job_id)
    if job is None:
        raise HTTPException(404, "Job not found")
    if job["status"] != "done":
        raise HTTPException(409, "Translation not yet complete")
    result_path = job["result_file"]
    if not result_path or not os.path.exists(result_path):
        raise HTTPException(410, "Result file has been cleaned up")
    filename = f"translated_{job['language']}.docx"
    return FileResponse(result_path, filename=filename, media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document")

@app.get("/api/providers")
def api_providers():
    providers = []
    for name, cfg in config.get("providers", {}).items():
        providers.append({"name": name, "model": cfg.get("model", "")})
    return {"default_provider": config.get("default_provider", ""), "providers": providers}

# Serve frontend in production
frontend_dist = Path(__file__).parent / "frontend" / "dist"
if frontend_dist.exists():
    app.mount("/", StaticFiles(directory=str(frontend_dist), html=True), name="frontend")

# --- Tests ---

def test_create_job():
    from fastapi.testclient import TestClient
    from docx import Document
    import tempfile
    import os

    doc = Document()
    doc.add_paragraph("Test")
    tmp = os.path.join(tempfile.mkdtemp(), "test.docx")
    doc.save(tmp)

    with TestClient(app) as client:
        with open(tmp, "rb") as f:
            resp = client.post(
                "/api/translate",
                files={"file": ("test.docx", f, "application/vnd.openxmlformats-officedocument.wordprocessingml.document")},
                data={"lang": "ro"},
            )
        assert resp.status_code == 201
        data = resp.json()
        assert "job_id" in data
        assert data["status"] == "pending"
    os.unlink(tmp)

def test_status_not_found():
    from fastapi.testclient import TestClient
    with TestClient(app) as client:
        resp = client.get("/api/translate/nonexistent-id/status")
        assert resp.status_code == 404

def test_download_not_done():
    from fastapi.testclient import TestClient
    with TestClient(app) as client:
        resp = client.get("/api/translate/nonexistent-id/download")
        assert resp.status_code == 404

def test_providers():
    from fastapi.testclient import TestClient
    with TestClient(app) as client:
        resp = client.get("/api/providers")
        assert resp.status_code == 200
        data = resp.json()
        assert "default_provider" in data
        assert "providers" in data
