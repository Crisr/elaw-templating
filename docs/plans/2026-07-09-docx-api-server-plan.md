# DOCX Translation API Server — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Turn the existing DOCX CLI translator into an async API server with a React frontend, containerized with Docker.

**Architecture:** FastAPI server with SQLite job queue. Background worker thread processes translations with progress tracking. React frontend (Vite + Tailwind) with drag-and-drop upload, polling progress bar, and download. Multi-stage Docker build (node → python).

**Tech Stack:** Python 3.11+, FastAPI, uvicorn, SQLite3, react, vite, tailwindcss, Docker

## Global Constraints

- `translate_docx.py` must remain unmodified in its CLI interface and behavior
- Only `translate_all()` gets a backward-compatible optional `progress_callback` param added
- All provider config stays in `config.json`, read at server startup
- Languages: `ro` and `en` only
- Jobs in SQLite, files on disk with rolling cleanup (max 10 pairs)
- Frontend is pure React + Vite + Tailwind, no SSR, no routing library
- Single Docker container for production

---

### Task 1: Update Dependencies

**Files:**
- Modify: `requirements.txt`

**Interfaces:**
- Consumes: existing `requirements.txt`
- Produces: updated `requirements.txt` with FastAPI, uvicorn, python-multipart

- [ ] **Step 1: Update requirements.txt**

Append to the existing file:

```
fastapi>=0.115.0
uvicorn[standard]>=0.32.0
python-multipart>=0.0.18
```

--- Start of manual proofread area ---

--- End of manual proofread area ---

- [ ] **Step 2: Install and verify**

```bash
cd /Users/cristianradu/Documents/programming/elaw-templating && pip install -r requirements.txt
```

Expected: All packages install cleanly.

- [ ] **Step 3: Commit**

```bash
git add requirements.txt
git commit -m "chore: add fastapi, uvicorn, python-multipart deps"
```

---

### Task 2: Add Progress Callback to translate_all

**Files:**
- Modify: `translate_docx.py` — add `progress_callback` parameter to `translate_all()`

**Interfaces:**
- Consumes: existing `translate_all(paragraphs, target_lang, provider, concurrency=0)`
- Produces: `translate_all(paragraphs, target_lang, provider, concurrency=0, progress_callback=None)`
  - `progress_callback(done: int, total: int) -> None` — called after each chunk completes

- [ ] **Step 1: Modify `translate_all` signature**

In `translate_docx.py`, change the function signature to add the callback:

```python
def translate_all(paragraphs, target_lang, provider, concurrency=0, progress_callback=None):
```

- [ ] **Step 2: Replace `_show_progress` calls with callback**

Replace the two `_show_progress` calls inside `translate_all`:

```python
# Line 297: _show_progress(0, total)
if progress_callback:
    progress_callback(0, total)
else:
    _show_progress(0, total)
```

```python
# Line 309: _show_progress(done, total)
if progress_callback:
    progress_callback(done, total)
else:
    _show_progress(done, total)
```

- [ ] **Step 3: Run existing tests to confirm backward compatibility**

```bash
cd /Users/cristianradu/Documents/programming/elaw-templating && python3 translate_docx.py
```

Expected: All tests PASS. The CLI still works identically.

- [ ] **Step 4: Commit**

```bash
git add translate_docx.py
git commit -m "feat: add optional progress_callback to translate_all"
```

---

### Task 3: Create server.py — FastAPI App + SQLite Job Manager

**Files:**
- Create: `server.py`

**Interfaces:**
- Consumes: `translate_docx.load_config()`, `translate_docx.get_provider()`, `worker.enqueue_job()`
- Produces: FastAPI app with routes:
  - `POST /api/translate` → `{job_id, status}`
  - `GET /api/translate/{job_id}/status` → `{job_id, status, progress, total, error}`
  - `GET /api/translate/{job_id}/download` → file stream
  - `GET /api/providers` → `{default_provider, providers}`
  - `GET /` → serves frontend static files

- [ ] **Step 1: Write the failing server test**

Create tests at the bottom of `server.py` using FastAPI's `TestClient`:

```python
def test_create_job():
    from fastapi.testclient import TestClient
    import io
    from docx import Document
    import tempfile
    import os

    # Create a minimal DOCX
    doc = Document()
    doc.add_paragraph("Test document content")
    tmp = os.path.join(tempfile.mkdtemp(), "test.docx")
    doc.save(tmp)

    client = TestClient(app)
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
```

- [ ] **Step 2: Write the minimal server.py**

```python
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

    client = TestClient(app)
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
    client = TestClient(app)
    resp = client.get("/api/translate/nonexistent-id/status")
    assert resp.status_code == 404

def test_download_not_done():
    from fastapi.testclient import TestClient
    client = TestClient(app)
    resp = client.get("/api/translate/nonexistent-id/download")
    assert resp.status_code == 404

def test_providers():
    from fastapi.testclient import TestClient
    client = TestClient(app)
    resp = client.get("/api/providers")
    assert resp.status_code == 200
    data = resp.json()
    assert "default_provider" in data
    assert "providers" in data
```

- [ ] **Step 3: Run the failing tests**

```bash
cd /Users/cristianradu/Documents/programming/elaw-templating && python3 -c "
import sys
sys.path.insert(0, '.')
from server import test_create_job, test_status_not_found, test_download_not_done, test_providers
test_create_job()
print('test_create_job PASS')
test_status_not_found()
print('test_status_not_found PASS')
test_download_not_done()
print('test_download_not_done PASS')
test_providers()
print('test_providers PASS')
"
```

Note: The first run will fail because `worker.py` doesn't exist yet. Create a minimal stub:

```python
class TranslationWorker:
    def __init__(self):
        self._queue = []
        self._running = False
    def start(self): pass
    def stop(self): pass
    def enqueue(self, job_id): pass
```

Run tests again — should pass for the route-level tests (worker is a no-op stub).

- [ ] **Step 4: Commit**

```bash
git add server.py
git commit -m "feat: add FastAPI server with SQLite job manager and API routes"
```

---

### Task 4: Create worker.py — Background Translation Worker

**Files:**
- Create: `worker.py`

**Interfaces:**
- Consumes: `server.create_job()`, `server.update_job()`, `server.get_job()`, `server.enforce_file_limit()`, `translate_docx`
- Produces: `TranslationWorker` class with `start()`, `stop()`, `enqueue(job_id)`
  - Worker polls queue, picks pending jobs, runs `translate_all()` with progress callback

- [ ] **Step 1: Write worker.py**

```python
import threading
import time
import json
import os
import shutil
from pathlib import Path

import translate_docx
from server import update_job, get_job, enforce_file_limit, config, UPLOAD_DIR


class TranslationWorker:
    def __init__(self):
        self._queue = []
        self._lock = threading.Lock()
        self._cv = threading.Condition(self._lock)
        self._running = False
        self._thread = None

    def start(self):
        self._running = True
        self._thread = threading.Thread(target=self._run, daemon=True)
        self._thread.start()

    def stop(self):
        self._running = False
        with self._cv:
            self._cv.notify_all()
        if self._thread:
            self._thread.join(timeout=5)

    def enqueue(self, job_id):
        with self._cv:
            self._queue.append(job_id)
            self._cv.notify()

    def _run(self):
        while self._running:
            job_id = None
            with self._cv:
                while self._running and not self._queue:
                    self._cv.wait(timeout=1)
                if self._queue:
                    job_id = self._queue.pop(0)
            if job_id and self._running:
                self._process(job_id)

    def _process(self, job_id):
        try:
            job = get_job(job_id)
            if job is None:
                return

            source_path = job["source_file"]
            lang = "Romanian" if job["language"] == "ro" else "English"
            mode = job.get("mode", "inline")
            provider_name = job.get("provider")
            model_override = job.get("model")

            provider = translate_docx.get_provider(config, provider_name)
            if model_override:
                provider["model"] = model_override

            update_job(job_id, status="running", progress=0)

            def progress_callback(done, total):
                update_job(job_id, progress=done, total=total)

            paragraphs = translate_docx.extract_paragraphs(source_path)
            originals = __import__("copy").deepcopy(paragraphs) if mode == "side-by-side" else None
            translated = translate_docx.translate_all(
                paragraphs, lang, provider, progress_callback=progress_callback
            )

            result_path = UPLOAD_DIR / f"{job_id}_result.docx"
            if mode == "side-by-side":
                translate_docx.write_side_by_side(source_path, originals, translated, str(result_path))
            else:
                translate_docx.write_inline(source_path, translated, str(result_path))

            update_job(job_id, status="done", progress=100, result_file=str(result_path))
            enforce_file_limit()

        except Exception as e:
            update_job(job_id, status="failed", error=str(e))
```

- [ ] **Step 2: Test the worker processes a job**

Since this requires an LLM call, the test is integration-level. For unit testing, verify the worker's job lifecycle:

```python
def test_worker_enqueue_dequeue():
    import server
    # Create a job directly
    job_id = server.create_job("ro", "inline", None, None)
    from worker import TranslationWorker
    w = TranslationWorker()
    w.start()
    w.enqueue(job_id)
    # Worker should pick up and immediately fail (no source file)
    time.sleep(0.5)
    job = server.get_job(job_id)
    assert job["status"] == "failed"  # no source file exists
    w.stop()
```

- [ ] **Step 3: Commit**

```bash
git add worker.py
git commit -m "feat: add background translation worker with progress callbacks"
```

---

### Task 5: Scaffold React Frontend with Vite + Tailwind

**Files:**
- Create: `frontend/package.json`
- Create: `frontend/vite.config.ts`
- Create: `frontend/tsconfig.json`
- Create: `frontend/tsconfig.app.json`
- Create: `frontend/tsconfig.node.json`
- Create: `frontend/tailwind.config.js`
- Create: `frontend/postcss.config.js`
- Create: `frontend/index.html`
- Create: `frontend/src/main.tsx`
- Create: `frontend/src/index.css`
- Create: `frontend/src/App.tsx`

- [ ] **Step 1: Create package.json**

```json
{
  "name": "docx-translator-frontend",
  "private": true,
  "version": "0.1.0",
  "type": "module",
  "scripts": {
    "dev": "vite",
    "build": "tsc -b && vite build",
    "preview": "vite preview"
  },
  "dependencies": {
    "react": "^18.3.1",
    "react-dom": "^18.3.1"
  },
  "devDependencies": {
    "@types/react": "^18.3.12",
    "@types/react-dom": "^18.3.1",
    "@vitejs/plugin-react": "^4.3.4",
    "autoprefixer": "^10.4.20",
    "postcss": "^8.4.49",
    "tailwindcss": "^3.4.17",
    "typescript": "^5.6.3",
    "vite": "^6.0.5"
  }
}
```

- [ ] **Step 2: Create vite.config.ts**

```typescript
import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

export default defineConfig({
  plugins: [react()],
  server: {
    proxy: {
      '/api': 'http://localhost:8000',
    },
  },
})
```

- [ ] **Step 3: Create tsconfig files**

`tsconfig.json`:
```json
{
  "files": [],
  "references": [
    { "path": "./tsconfig.app.json" },
    { "path": "./tsconfig.node.json" }
  ]
}
```

`tsconfig.app.json`:
```json
{
  "compilerOptions": {
    "target": "ES2020",
    "useDefineForClassFields": true,
    "lib": ["ES2020", "DOM", "DOM.Iterable"],
    "module": "ESNext",
    "skipLibCheck": true,
    "moduleResolution": "bundler",
    "allowImportingTsExtensions": true,
    "isolatedModules": true,
    "moduleDetection": "force",
    "noEmit": true,
    "jsx": "react-jsx",
    "strict": true,
    "noUnusedLocals": true,
    "noUnusedParameters": true,
    "noFallthroughCasesInSwitch": true,
    "noUncheckedSideEffectImports": true
  },
  "include": ["src"]
}
```

`tsconfig.node.json`:
```json
{
  "compilerOptions": {
    "target": "ES2022",
    "lib": ["ES2023"],
    "module": "ESNext",
    "skipLibCheck": true,
    "moduleResolution": "bundler",
    "allowImportingTsExtensions": true,
    "isolatedModules": true,
    "moduleDetection": "force",
    "noEmit": true,
    "strict": true,
    "noUnusedLocals": true,
    "noUnusedParameters": true,
    "noFallthroughCasesInSwitch": true,
    "noUncheckedSideEffectImports": true
  },
  "include": ["vite.config.ts"]
}
```

- [ ] **Step 4: Create tailwind.config.js**

```javascript
/** @type {import('tailwindcss').Config} */
export default {
  content: ['./index.html', './src/**/*.{js,ts,jsx,tsx}'],
  theme: {
    extend: {},
  },
  plugins: [],
}
```

- [ ] **Step 5: Create postcss.config.js**

```javascript
export default {
  plugins: {
    tailwindcss: {},
    autoprefixer: {},
  },
}
```

- [ ] **Step 6: Create index.html**

```html
<!DOCTYPE html>
<html lang="en">
  <head>
    <meta charset="UTF-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0" />
    <title>DOCX Translator</title>
  </head>
  <body class="bg-gray-50 min-h-screen">
    <div id="root"></div>
    <script type="module" src="/src/main.tsx"></script>
  </body>
</html>
```

- [ ] **Step 7: Create src/main.tsx**

```typescript
import { StrictMode } from 'react'
import { createRoot } from 'react-dom/client'
import './index.css'
import App from './App'

createRoot(document.getElementById('root')!).render(
  <StrictMode>
    <App />
  </StrictMode>,
)
```

- [ ] **Step 8: Create src/index.css**

```css
@tailwind base;
@tailwind components;
@tailwind utilities;
```

- [ ] **Step 9: Create src/App.tsx** (placeholder)

```typescript
function App() {
  return (
    <div className="max-w-2xl mx-auto p-8">
      <h1 className="text-2xl font-bold mb-6">DOCX Translator</h1>
      <p>Frontend scaffolded. Components coming in next task.</p>
    </div>
  )
}

export default App
```

- [ ] **Step 10: Install and verify build**

```bash
cd /Users/cristianradu/Documents/programming/elaw-templating/frontend && npm install && npm run build
```

Expected: Build succeeds, `frontend/dist/` is created with `index.html`.

- [ ] **Step 11: Commit**

```bash
git add frontend/
git commit -m "feat: scaffold React + Vite + Tailwind frontend"
```

---

### Task 6: Build Frontend Components

**Files:**
- Modify: `frontend/src/App.tsx`
- Create: `frontend/src/components/DropZone.tsx`
- Create: `frontend/src/components/OptionsForm.tsx`
- Create: `frontend/src/components/ProgressBar.tsx`
- Create: `frontend/src/components/DownloadLink.tsx`

**Interfaces:**
- Consumes: `POST /api/translate`, `GET /api/translate/{id}/status`, `GET /api/translate/{id}/download`, `GET /api/providers`
- Produces: Full UI with drag-and-drop, options, progress, download

- [ ] **Step 1: Create App.tsx (full)**

```typescript
import { useState, useCallback } from 'react'
import DropZone from './components/DropZone'
import OptionsForm from './components/OptionsForm'
import ProgressBar from './components/ProgressBar'
import DownloadLink from './components/DownloadLink'

type Status = 'idle' | 'uploading' | 'translating' | 'done' | 'error'

function App() {
  const [file, setFile] = useState<File | null>(null)
  const [lang, setLang] = useState('ro')
  const [mode, setMode] = useState('inline')
  const [providerName, setProviderName] = useState('')
  const [modelName, setModelName] = useState('')
  const [status, setStatus] = useState<Status>('idle')
  const [jobId, setJobId] = useState<string | null>(null)
  const [progress, setProgress] = useState(0)
  const [total, setTotal] = useState(0)
  const [error, setError] = useState('')

  const handleSubmit = useCallback(async () => {
    if (!file) return
    setStatus('uploading')
    setError('')

    const formData = new FormData()
    formData.append('file', file)
    formData.append('lang', lang)
    formData.append('mode', mode)
    if (providerName) formData.append('provider', providerName)
    if (modelName) formData.append('model', modelName)

    try {
      const resp = await fetch('/api/translate', { method: 'POST', body: formData })
      if (!resp.ok) throw new Error(`Upload failed: ${resp.statusText}`)
      const data = await resp.json()
      setJobId(data.job_id)
      setStatus('translating')

      // Poll status
      const poll = setInterval(async () => {
        const sResp = await fetch(`/api/translate/${data.job_id}/status`)
        if (!sResp.ok) {
          clearInterval(poll)
          setStatus('error')
          setError('Failed to fetch status')
          return
        }
        const sData = await sResp.json()
        setProgress(sData.progress)
        setTotal(sData.total)

        if (sData.status === 'done') {
          clearInterval(poll)
          setStatus('done')
        } else if (sData.status === 'failed') {
          clearInterval(poll)
          setStatus('error')
          setError(sData.error || 'Translation failed')
        }
      }, 500)
    } catch (e) {
      setStatus('error')
      setError(e instanceof Error ? e.message : 'Unknown error')
    }
  }, [file, lang, mode, providerName, modelName])

  const handleReset = () => {
    setFile(null)
    setJobId(null)
    setStatus('idle')
    setProgress(0)
    setTotal(0)
    setError('')
  }

  return (
    <div className="min-h-screen bg-gradient-to-br from-blue-50 to-indigo-100 flex items-start justify-center p-4 sm:p-8">
      <div className="w-full max-w-xl bg-white rounded-2xl shadow-lg p-6 sm:p-8 space-y-6">
        <h1 className="text-2xl font-bold text-gray-800">DOCX Translator</h1>

        {status === 'idle' || status === 'uploading' ? (
          <>
            <DropZone file={file} onFile={setFile} disabled={status === 'uploading'} />
            <OptionsForm
              lang={lang}
              mode={mode}
              providerName={providerName}
              modelName={modelName}
              onLangChange={setLang}
              onModeChange={setMode}
              onProviderChange={setProviderName}
              onModelChange={setModelName}
            />
            <button
              onClick={handleSubmit}
              disabled={!file || status === 'uploading'}
              className="w-full py-3 px-4 bg-indigo-600 text-white font-medium rounded-xl disabled:opacity-40 disabled:cursor-not-allowed hover:bg-indigo-700 transition-colors"
            >
              {status === 'uploading' ? 'Uploading...' : 'Convert'}
            </button>
          </>
        ) : status === 'translating' ? (
          <ProgressBar progress={progress} total={total} />
        ) : status === 'done' && jobId ? (
          <DownloadLink jobId={jobId} onReset={handleReset} />
        ) : status === 'error' ? (
          <div className="space-y-4">
            <div className="bg-red-50 border border-red-200 text-red-700 rounded-xl p-4">
              {error || 'An error occurred'}
            </div>
            <button
              onClick={handleReset}
              className="w-full py-3 px-4 bg-gray-600 text-white font-medium rounded-xl hover:bg-gray-700 transition-colors"
            >
              Try Again
            </button>
          </div>
        ) : null}
      </div>
    </div>
  )
}

export default App
```

- [ ] **Step 2: Create DropZone.tsx**

```typescript
import { useCallback, useRef, useState } from 'react'

interface Props {
  file: File | null
  onFile: (f: File) => void
  disabled: boolean
}

export default function DropZone({ file, onFile, disabled }: Props) {
  const [dragging, setDragging] = useState(false)
  const inputRef = useRef<HTMLInputElement>(null)

  const handleDrop = useCallback(
    (e: React.DragEvent) => {
      e.preventDefault()
      setDragging(false)
      if (disabled) return
      const f = e.dataTransfer.files[0]
      if (f && f.name.endsWith('.docx')) onFile(f)
    },
    [disabled, onFile]
  )

  const handleChange = useCallback(
    (e: React.ChangeEvent<HTMLInputElement>) => {
      const f = e.target.files?.[0]
      if (f) onFile(f)
    },
    [onFile]
  )

  return (
    <div
      onDragOver={(e) => { e.preventDefault(); if (!disabled) setDragging(true) }}
      onDragLeave={() => setDragging(false)}
      onDrop={handleDrop}
      onClick={() => inputRef.current?.click()}
      className={`border-2 border-dashed rounded-xl p-8 text-center cursor-pointer transition-colors ${
        dragging
          ? 'border-indigo-500 bg-indigo-50'
          : 'border-gray-300 hover:border-gray-400 bg-gray-50'
      } ${disabled ? 'opacity-50 cursor-not-allowed' : ''}`}
    >
      <input
        ref={inputRef}
        type="file"
        accept=".docx"
        className="hidden"
        onChange={handleChange}
      />
      {file ? (
        <p className="text-gray-700 font-medium">{file.name}</p>
      ) : (
        <div>
          <p className="text-gray-500">Drag & drop a .docx file here</p>
          <p className="text-gray-400 text-sm mt-1">or click to browse</p>
        </div>
      )}
    </div>
  )
}
```

- [ ] **Step 3: Create OptionsForm.tsx**

```typescript
import { useState, useEffect } from 'react'

interface Props {
  lang: string
  mode: string
  providerName: string
  modelName: string
  onLangChange: (v: string) => void
  onModeChange: (v: string) => void
  onProviderChange: (v: string) => void
  onModelChange: (v: string) => void
}

interface Provider {
  name: string
  model: string
}

export default function OptionsForm({
  lang, mode, providerName, modelName,
  onLangChange, onModeChange, onProviderChange, onModelChange,
}: Props) {
  const [providers, setProviders] = useState<Provider[]>([])
  const [defaultProvider, setDefaultProvider] = useState('')

  useEffect(() => {
    fetch('/api/providers')
      .then((r) => r.json())
      .then((data) => {
        setProviders(data.providers || [])
        setDefaultProvider(data.default_provider || '')
        if (!providerName && data.default_provider) {
          onProviderChange(data.default_provider)
        }
      })
      .catch(() => {})
  }, []) // eslint-disable-line react-hooks/exhaustive-deps

  const selectedProvider = providers.find((p) => p.name === (providerName || defaultProvider))

  return (
    <div className="space-y-4">
      <div className="grid grid-cols-2 gap-4">
        <div>
          <label className="block text-sm font-medium text-gray-700 mb-1">Language</label>
          <select
            value={lang}
            onChange={(e) => onLangChange(e.target.value)}
            className="w-full border border-gray-300 rounded-lg px-3 py-2 text-sm focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500"
          >
            <option value="ro">Romanian</option>
            <option value="en">English</option>
          </select>
        </div>
        <div>
          <label className="block text-sm font-medium text-gray-700 mb-1">Mode</label>
          <select
            value={mode}
            onChange={(e) => onModeChange(e.target.value)}
            className="w-full border border-gray-300 rounded-lg px-3 py-2 text-sm focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500"
          >
            <option value="inline">Inline</option>
            <option value="side-by-side">Side-by-Side</option>
          </select>
        </div>
      </div>
      <div className="grid grid-cols-2 gap-4">
        <div>
          <label className="block text-sm font-medium text-gray-700 mb-1">Provider</label>
          <select
            value={providerName || defaultProvider}
            onChange={(e) => {
              onProviderChange(e.target.value)
              const p = providers.find((pr) => pr.name === e.target.value)
              if (p) onModelChange(p.model)
            }}
            className="w-full border border-gray-300 rounded-lg px-3 py-2 text-sm focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500"
          >
            {providers.map((p) => (
              <option key={p.name} value={p.name}>{p.name}</option>
            ))}
          </select>
        </div>
        <div>
          <label className="block text-sm font-medium text-gray-700 mb-1">Model</label>
          <input
            type="text"
            value={modelName || selectedProvider?.model || ''}
            onChange={(e) => onModelChange(e.target.value)}
            placeholder={selectedProvider?.model || 'Model name'}
            className="w-full border border-gray-300 rounded-lg px-3 py-2 text-sm focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500"
          />
        </div>
      </div>
    </div>
  )
}
```

- [ ] **Step 4: Create ProgressBar.tsx**

```typescript
interface Props {
  progress: number
  total: number
}

export default function ProgressBar({ progress, total }: Props) {
  const pct = total > 0 ? Math.round((progress / total) * 100) : 0

  return (
    <div className="space-y-3">
      <p className="text-sm text-gray-600 text-center">
        Translating... {progress} / {total} chunks ({pct}%)
      </p>
      <div className="w-full bg-gray-200 rounded-full h-4 overflow-hidden">
        <div
          className="h-full bg-indigo-600 rounded-full transition-all duration-300 ease-out"
          style={{ width: `${pct}%` }}
        />
      </div>
    </div>
  )
}
```

- [ ] **Step 5: Create DownloadLink.tsx**

```typescript
interface Props {
  jobId: string
  onReset: () => void
}

export default function DownloadLink({ jobId, onReset }: Props) {
  return (
    <div className="text-center space-y-4">
      <div className="bg-green-50 border border-green-200 text-green-700 rounded-xl p-4">
        Translation complete!
      </div>
      <a
        href={`/api/translate/${jobId}/download`}
        className="inline-block w-full py-3 px-4 bg-indigo-600 text-white font-medium rounded-xl text-center hover:bg-indigo-700 transition-colors"
      >
        Download Translated File
      </a>
      <button
        onClick={onReset}
        className="w-full py-2 px-4 text-gray-600 hover:text-gray-800 text-sm transition-colors"
      >
        Translate Another
      </button>
    </div>
  )
}
```

- [ ] **Step 6: Verify build**

```bash
cd /Users/cristianradu/Documents/programming/elaw-templating/frontend && npm run build
```

Expected: Build succeeds, `frontend/dist/` contains the compiled app.

- [ ] **Step 7: Commit**

```bash
git add frontend/src/
git commit -m "feat: add frontend components - DropZone, OptionsForm, ProgressBar, DownloadLink"
```

---

### Task 7: Docker Configuration

**Files:**
- Create: `Dockerfile`
- Create: `.dockerignore`

- [ ] **Step 1: Create Dockerfile**

```dockerfile
# Stage 1: Build frontend
FROM node:20-alpine AS frontend-builder

WORKDIR /app/frontend
COPY frontend/package.json frontend/package-lock.json ./
RUN npm ci
COPY frontend/ .
RUN npm run build

# Stage 2: Python runtime
FROM python:3.11-slim

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY translate_docx.py messages.py server.py worker.py config.json ./
COPY --from=frontend-builder /app/frontend/dist ./frontend/dist

RUN mkdir -p uploads

EXPOSE 8000

CMD ["uvicorn", "server:app", "--host", "0.0.0.0", "--port", "8000"]
```

Wait for `frontend/package-lock.json` to exist (created by `npm install` in task 5). If not, remove the lockfile line from the Dockerfile or generate it first.

- [ ] **Step 2: Generate package-lock.json**

```bash
cd /Users/cristianradu/Documents/programming/elaw-templating/frontend && npm install
```

This creates `package-lock.json`.

- [ ] **Step 3: Create .dockerignore**

```
__pycache__
*.pyc
.env
.env.*
secrets.json
.git
.gitignore
**/node_modules
uploads/
jobs.db
```

- [ ] **Step 4: Commit**

```bash
git add Dockerfile .dockerignore frontend/package-lock.json
git commit -m "feat: add Docker multi-stage build and .dockerignore"
```

---

### Task 8: Integration Smoke Test

- [ ] **Step 1: Build and run with Docker**

```bash
cd /Users/cristianradu/Documents/programming/elaw-templating && docker build -t docx-translator .
```

- [ ] **Step 2: Run the container**

```bash
docker run -d -p 8000:8000 -v "$(pwd)/uploads:/app/uploads" -v "$(pwd)/jobs.db:/app/jobs.db" --name docx-translator docx-translator
```

- [ ] **Step 3: Test the API health**

```bash
# Check providers endpoint
curl -s http://localhost:8000/api/providers | python3 -m json.tool

# Check frontend is served
curl -s -o /dev/null -w "%{http_code}" http://localhost:8000/
```

Expected: Both return 200.

- [ ] **Step 4: Test a translation job (requires working provider)**

```bash
# Create a test docx
python3 -c "
from docx import Document
doc = Document()
doc.add_paragraph('Hello world, this is a test.')
doc.add_paragraph('Second paragraph with {{placeholder}} content.')
doc.save('/tmp/test_translate.docx')
"

# Upload and translate
JOB_RESP=$(curl -s -X POST http://localhost:8000/api/translate \
  -F "file=@/tmp/test_translate.docx" \
  -F "lang=ro")
echo "$JOB_RESP"
JOB_ID=$(echo "$JOB_RESP" | python3 -c "import sys,json; print(json.load(sys.stdin)['job_id'])")

# Poll until done
for i in $(seq 1 60); do
  STATUS=$(curl -s http://localhost:8000/api/translate/$JOB_ID/status)
  STATE=$(echo "$STATUS" | python3 -c "import sys,json; print(json.load(sys.stdin)['status'])")
  echo "Poll $i: $STATE"
  if [ "$STATE" = "done" ] || [ "$STATE" = "failed" ]; then
    break
  fi
  sleep 2
done

# Download if done
if [ "$STATE" = "done" ]; then
  curl -s -o /tmp/translated.docx http://localhost:8000/api/translate/$JOB_ID/download
  echo "Downloaded to /tmp/translated.docx"
fi
```

- [ ] **Step 5: Stop the container**

```bash
docker stop docx-translator && docker rm docx-translator
```

---

### Task 9: Self-Review

1. **Spec coverage:** All spec sections covered — API endpoints (Task 3), SQLite jobs (Task 3), worker (Task 4), frontend components (Task 6), Docker (Task 7), file cleanup (Task 3 via `enforce_file_limit`), providers endpoint (Task 3).

2. **Placeholder scan:** No TBD, TODO, or incomplete steps. All code blocks are complete.

3. **Type consistency:** `translate_all()` signature matches between Task 2 (add `progress_callback`) and Task 4 (worker calls it). `create_job()`, `update_job()`, `get_job()` signatures consistent between Task 3 and Task 4.

4. **No gaps:** Everything from the spec has a corresponding task.
