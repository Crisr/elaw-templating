# DOCX Translation API Server — Design Spec

**Date:** 2026-07-09
**Status:** Design

## 1. Architecture

```
Frontend (React + Vite + Tailwind)          Backend (FastAPI, single process)
┌──────────────────────┐     ┌─────────────────────────────────────────────┐
│  DropZone            │────▶│  POST /api/translate (multipart)            │
│  OptionsForm         │     │  GET  /api/translate/{id}/status            │
│  ProgressBar         │◀────│  GET  /api/translate/{id}/download          │
│  DownloadLink        │     │  GET  /api/providers                        │
│                      │     │  GET  / (serves frontend/dist/)             │
└──────────────────────┘     │                                             │
                             │  ┌──────────┐  ┌────────────┐  ┌─────────┐ │
                             │  │   Job    │──│   Worker   │──│ SQLite  │ │
                             │  │  Manager │  │  (thread)  │  │ jobs.db │ │
                             │  └──────────┘  │  imports   │  └─────────┘ │
                             │                │  translate_ │             │
                             │                │  docx.py    │  uploads/    │
                             │                └────────────┘  .docx files │
                             └─────────────────────────────────────────────┘
```

- Core translation logic in `translate_docx.py` is **unchanged** — the server imports and calls it
- CLI (`python translate_docx.py input.docx --lang ro`) continues to work independently

## 2. API Endpoints

| Method | Path | Description |
|--------|------|-------------|
| `POST` | `/api/translate` | Upload DOCX + options → `{ job_id, status }` |
| `GET` | `/api/translate/{id}/status` | Poll job → `{ status, progress, total, error }` |
| `GET` | `/api/translate/{id}/download` | Download result DOCX |
| `GET` | `/api/providers` | List providers from config.json |
| `GET` | `/` | Serve React frontend |

### POST /api/translate

```
Content-Type: multipart/form-data

Fields:
  file      (file, required)    — .docx file to translate
  lang      (string, required)  — "ro" or "en"
  mode      (string, default: "inline") — "inline" or "side-by-side"
  provider  (string, optional)  — provider key from config.json
  model     (string, optional)  — override model name

Response 201:
{
  "job_id": "uuid-string",
  "status": "pending"
}
```

### GET /api/translate/{job_id}/status

```
Response 200:
{
  "job_id": "uuid-string",
  "status": "running",
  "progress": 7,
  "total": 15,
  "error": null
}
```

### GET /api/translate/{job_id}/download

```
Response 200: application/vnd.openxmlformats-officedocument.wordprocessingml.document
Content-Disposition: attachment; filename="document_ro.docx"
```

Returns 404 if job doesn't exist, 409 if not done, 410 if file cleaned up.

### GET /api/providers

```
Response 200:
{
  "default_provider": "local",
  "providers": [
    {"name": "local", "model": "Carnice 9B"},
    {"name": "openai", "model": "gpt-4o-mini"}
  ]
}
```

## 3. Job Queue (SQLite)

```sql
CREATE TABLE jobs (
    id TEXT PRIMARY KEY,
    status TEXT DEFAULT 'pending',  -- pending | running | done | failed
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
);
```

- Thread-safe with `threading.Lock`
- Worker dequeues `pending` jobs, sets `running`, runs `translate_all()` with a progress callback
- On completion: writes result DOCX, sets `done`; on failure: sets `failed` with error message

## 4. File Management

- `uploads/{job_id}_source.docx` — uploaded original
- `uploads/{job_id}_result.docx` — translated output
- Rolling cleanup: keep max 10 pairs (20 files). After each job completion, if total files > 20, delete oldest by mtime until at or under limit.
- Job metadata row stays in SQLite even after file cleanup (download returns 410 Gone)

## 5. Frontend (React + Vite + Tailwind)

**Components:**
- `DropZone.tsx` — drag-and-drop area with click-to-browse fallback
- `OptionsForm.tsx` — language (ro/en), mode (inline/side-by-side), provider dropdown, model text input
- `ProgressBar.tsx` — CSS progress bar with "N / M chunks (X%)" label, polls `/status` every 500ms
- `DownloadLink.tsx` — shown on completion, links to `/download`

**States:** idle → uploading → translating (progress) → done/error

## 6. Development vs Production

| | Dev | Prod (Docker) |
|---|---|---|
| Frontend | `npm run dev` (Vite, port 5173) | Built to `frontend/dist/` |
| Backend | `uvicorn server:app` (port 8000) | Same, serves static from `dist/` |
| CORS | Allowed from `localhost:5173` | Not needed (same origin) |

## 7. Docker (Multi-stage Build)

Stage 1 (node): `npm ci && npm run build` → `frontend/dist/`
Stage 2 (python): `pip install`, copy `dist/`, run uvicorn

## 8. Desktop App Future

The `frontend/` directory is a standard Vite React app. It can be wrapped in Electron or Tauri with zero backend changes — just point the webview at the FastAPI server URL.

## 9. Out of Scope

- Authentication (local network only)
- Horizontal scaling (single process)
- File encryption at rest
- Document history beyond SQLite job metadata

## 10. Project Structure

```
elaw-templating/
├── translate_docx.py       # CLI — untouched
├── messages.py             # untouched
├── config.json             # untouched
├── requirements.txt        # +fastapi, uvicorn, python-multipart
├── server.py               # NEW: FastAPI routes, job manager, SQLite
├── worker.py               # NEW: Background translation worker
├── uploads/                # runtime
├── jobs.db                 # runtime
├── Dockerfile              # NEW: multi-stage build
├── train.json              # existing (for templating)
├── samples/                # existing
├── docs/
│   ├── specs/
│   └── plans/
└── frontend/               # NEW: React + Vite + Tailwind
    ├── package.json
    ├── vite.config.ts
    ├── tailwind.config.js
    ├── postcss.config.js
    ├── index.html
    ├── tsconfig.json
    ├── tsconfig.app.json
    ├── tsconfig.node.json
    └── src/
        ├── main.tsx
        ├── App.tsx
        ├── components/
        │   ├── DropZone.tsx
        │   ├── OptionsForm.tsx
        │   ├── ProgressBar.tsx
        │   └── DownloadLink.tsx
        └── index.css
```
