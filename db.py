import os
import uuid
import sqlite3
import threading
import time
from pathlib import Path

import translate_docx

UPLOAD_DIR = Path("uploads")
DB_PATH = "jobs.db"
MAX_FILE_PAIRS = 10

UPLOAD_DIR.mkdir(exist_ok=True)

VALID_COLUMNS = frozenset({
    "status", "progress", "total", "source_file", "result_file",
    "error", "language", "mode", "provider", "model",
})

_init_lock = threading.Lock()
_config = None
_config_lock = threading.Lock()


def get_config():
    global _config
    if _config is None:
        with _config_lock:
            if _config is None:
                _config = translate_docx.load_config()
    return _config


def reload_config():
    global _config
    with _config_lock:
        _config = translate_docx.load_config()


def _get_db():
    db = sqlite3.connect(DB_PATH, check_same_thread=False)
    db.row_factory = sqlite3.Row
    return db


def init_db():
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
    bad = set(kwargs) - VALID_COLUMNS
    if bad:
        raise ValueError(f"Invalid columns: {', '.join(sorted(bad))}")
    sets = ", ".join(f"{k} = ?" for k in kwargs)
    now = time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime())
    db = _get_db()
    db.execute(f"UPDATE jobs SET {sets}, updated_at = ? WHERE id = ?", list(kwargs.values()) + [now, job_id])
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
    files = sorted(UPLOAD_DIR.iterdir(), key=lambda f: f.stat().st_mtime)
    file_count = 0
    for f in files:
        if f.name.endswith("_source.docx") or f.name.endswith("_result.docx"):
            file_count += 1
    while file_count > MAX_FILE_PAIRS * 2 and files:
        f = files.pop(0)
        if f.name.endswith("_source.docx") or f.name.endswith("_result.docx"):
            f.unlink()
            file_count -= 1
