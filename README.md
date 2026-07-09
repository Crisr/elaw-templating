# DOCX Translator

Translate `.docx` files via a web UI or REST API. Built with FastAPI + React + SQLite.

## Quick Start (Docker)

```bash
docker build -t docx-translator .
docker run -d -p 8000:8000 --name docx-translator docx-translator
```

Open http://localhost:8000 in your browser.

## Running Without Docker

### Backend

```bash
pip install -r requirements.txt
python -m uvicorn server:app --host 0.0.0.0 --port 8000
```

### Frontend (Development)

```bash
cd frontend
npm install
npm run dev
```

In dev mode the frontend runs on port 5173 and proxies `/api` requests to the backend.

### Frontend Build

```bash
cd frontend
npm install
npm run build
```

The built files go to `frontend/dist/` and are served automatically by the FastAPI backend.

## Configuration

Edit `config.json` to add translation providers:

```json
{
  "default_provider": "local",
  "providers": {
    "local": { "model": "Carnice 9B" }
  }
}
```

## API Endpoints

| Method | Path | Description |
|--------|------|-------------|
| POST | `/api/translate` | Upload DOCX for translation |
| GET | `/api/translate/{id}/status` | Poll translation progress |
| GET | `/api/translate/{id}/download` | Download result |
| GET | `/api/providers` | List available providers |

## Translation Lookup

Translation messages are in `frontend/src/messages.ts`. Add a new locale by creating an object matching the `Messages` interface and updating the `messages` export.

## Project Structure

```
├── server.py              FastAPI app + API routes
├── worker.py              Background translation worker
├── db.py                  SQLite job persistence
├── translate_docx.py      Core translation logic
├── messages.py            Backend string constants
├── config.json            Provider configuration
├── requirements.txt       Python dependencies
├── Dockerfile             Multi-stage build
├── frontend/
│   ├── src/
│   │   ├── App.tsx        Main app component
│   │   ├── messages.ts    UI text (i18n-ready)
│   │   └── components/
│   │       ├── DropZone.tsx       File upload area
│   │       ├── OptionsForm.tsx    Language/mode/provider
│   │       ├── ProgressBar.tsx    Translation progress
│   │       └── DownloadLink.tsx   Result download
│   ├── public/
│   │   └── logo.png       Emplawra logo
│   ├── index.html
│   ├── tailwind.config.js
│   └── vite.config.ts
└── uploads/               Uploaded/result files (auto-created)
```
