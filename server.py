import os
import shutil
import time

from pathlib import Path
from contextlib import asynccontextmanager

import httpx

from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles
from typing import Optional

import db
import messages as msg


DOCX_MAGIC = b"PK\x03\x04"
_MODELS_CACHE = {}
_MODELS_CACHE_TTL = 30

async def _fetch_available_models(base_url):
    cached = _MODELS_CACHE.get(base_url)
    if cached and time.time() - cached["at"] < _MODELS_CACHE_TTL:
        return cached["models"]
    try:
        async with httpx.AsyncClient(timeout=3) as client:
            resp = await client.get(f"{base_url.rstrip('/')}/models")
            if resp.status_code == 200:
                data = resp.json()
                models = [m["id"] for m in data.get("data", [])]
                _MODELS_CACHE[base_url] = {"at": time.time(), "models": models}
                return models
    except Exception:
        pass
    return None


worker = None


@asynccontextmanager
async def lifespan(app: FastAPI):
    db.init_db()
    global worker
    from worker import TranslationWorker
    worker = TranslationWorker()
    worker.start()
    yield
    worker.stop()


app = FastAPI(lifespan=lifespan)


@app.post("/api/translate", status_code=201)
async def api_translate(
    file: UploadFile = File(...),
    lang: str = Form("ro"),
    mode: str = Form("inline"),
    transform2cell: bool = Form(False),
    provider: Optional[str] = Form(None),
    model: Optional[str] = Form(None),
):
    if transform2cell:
        job_id = db.create_job("", "transform2cell", provider, model)
    else:
        if lang not in ("ro", "en"):
            raise HTTPException(400, msg.MESSAGES["err_invalid_lang"])
        if mode not in ("inline", "side-by-side"):
            raise HTTPException(400, msg.MESSAGES["err_invalid_mode"])
        job_id = db.create_job(lang, mode, provider, model)

    head = await file.read(4)
    if head != DOCX_MAGIC:
        raise HTTPException(400, msg.MESSAGES["err_not_docx"])
    source_path = db.UPLOAD_DIR / f"{job_id}_source.docx"
    with open(source_path, "wb") as f:
        f.write(head)
        shutil.copyfileobj(file.file, f)

    db.update_job(job_id, source_file=str(source_path))
    worker.enqueue(job_id)

    return {"job_id": job_id, "status": "pending"}


@app.get("/api/translate/{job_id}/status")
def api_status(job_id: str):
    job = db.get_job(job_id)
    if job is None:
        raise HTTPException(404, msg.MESSAGES["err_job_not_found"])
    return {
        "job_id": job["id"],
        "status": job["status"],
        "progress": job["progress"],
        "total": job["total"],
        "error": job["error"],
    }


@app.get("/api/translate/{job_id}/download")
def api_download(job_id: str):
    job = db.get_job(job_id)
    if job is None:
        raise HTTPException(404, msg.MESSAGES["err_job_not_found"])
    if job["status"] != "done":
        raise HTTPException(409, msg.MESSAGES["err_not_done"])
    result_path = job["result_file"]
    if not result_path or not os.path.exists(result_path):
        raise HTTPException(410, msg.MESSAGES["err_result_cleaned"])
    filename = "transformed_cell.docx" if job["mode"] == "transform2cell" else f"translated_{job['language']}.docx"
    return FileResponse(result_path, filename=filename, media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document")


@app.get("/api/providers")
async def api_providers():
    config = db.get_config()
    providers = []
    for name, cfg in config.get("providers", {}).items():
        info = {"name": name, "model": cfg.get("model", "")}
        base_url = cfg.get("base_url", "")
        if base_url:
            models = await _fetch_available_models(base_url)
            if models:
                info["models"] = models
        providers.append(info)
    return {"default_provider": config.get("default_provider", ""), "providers": providers}


frontend_dist = Path(__file__).parent / "frontend" / "dist"
if frontend_dist.exists():
    app.mount("/", StaticFiles(directory=str(frontend_dist), html=True), name="frontend")
