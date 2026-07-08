# DOCX Translator — Technical Design Spec

**Date:** 2026-07-08
**Status:** Design

## Overview

A Python CLI script that translates `.docx` documents to Romanian and English using an LLM (OpenAI-compatible API), preserving all formatting. Supports normal inline replacement and a side-by-side original/translation layout.

## Architecture

```
DOCX → python-docx extractor → text chunks → LLM API → translated chunks → reconstructor → output DOCX
```

Single-file v1 (`translate_docx.py`), with a separate provider config.

## Provider Configuration (`config.json`)

Stored alongside the script. Each provider is an OpenAI-compatible endpoint:

```json
{
  "default_provider": "opencode-go",
  "providers": {
    "opencode-go": {
      "base_url": "https://api.opencode.ai/v1",
      "model": "opencode/deepseek-v4-flash",
      "api_key_env": "OPENCODE_API_KEY"
    },
    "lm-studio": {
      "base_url": "http://localhost:1234/v1",
      "model": "qwen-2.5-14b-instruct",
      "api_key_env": null
    },
    "openai": {
      "base_url": "https://api.openai.com/v1",
      "model": "gpt-4o-mini",
      "api_key_env": "OPENAI_API_KEY"
    }
  }
}
```

`api_key_env: null` means no API key sent (used for local servers). Keys loaded from environment or `.env` file.

## CLI Interface

```bash
python translate_docx.py input.docx --lang ro [options]
```

| Flag | Alias | Default | Description |
|------|-------|---------|-------------|
| `--lang` | `-l` | required | Target language: `ro` or `en` |
| `--mode` | `-m` | `inline` | `inline` or `side-by-side` |
| `--output` | `-o` | auto | Output path override |
| `--provider` | `-p` | from config | Provider key in config.json |
| `--model` | | provider default | Override model name |
| `--config` | `-c` | `./config.json` | Config file path |

Output auto-naming: `input_ro.docx` or `input_en.docx`.

## Chunking & Translation

### Extraction

Open the DOCX with `python-docx` and walk all paragraphs in the document body. Assign each paragraph a unique ID (`P0`, `P1`, ...). Record:

- Paragraph ID
- List of runs with their formatting (bold, italic, font, size, color)
- Paragraph-level formatting (alignment, spacing, indentation)
- Whether it's in a table cell (and which cell)

### Chunking

Group paragraphs into chunks of ~10 paragraphs (~2000 tokens to leave room for response). Chunks respect table boundaries — a table is never split across chunks to keep cell pairs together.

Each chunk is formatted for the LLM as:

```
[P0] First paragraph text
[P1] Second paragraph text with <b>bold</b> marker
[P2] Third paragraph text
```

### LLM Prompt

```
System: You are a legal document translator. Translate the following paragraphs
from the source language to {target_language}. Preserve all paragraph IDs exactly
([P0], [P1], ...). Preserve inline markup (bold, italic). Output ONLY the
translated paragraphs with their IDs — no greetings, no commentary, no extra text.

User: [P0] First paragraph
[P1] Second paragraph
...
```

### Retry Logic

- On HTTP timeout → retry up to 3 times with exponential backoff (1s, 2s, 4s)
- On malformed response (missing IDs) → fall back to sequential position matching
- On persistent failure → log the failing chunk and continue; report partial failure at end

## Formatting Preservation

### python-docx Run Model

Each `Paragraph` contains one or more `Run` objects. A run is a contiguous span of text with uniform formatting. The `Run` object holds:
- `text` — the string content
- `bold`, `italic`, `underline`, `font.size`, `font.color.rgb`, etc.

### Normal Mode (`--mode inline`)

Replace `run.text` in-place. All runs and their formatting remain untouched. Only the text content changes. The document structure (tables, headers, images) is completely preserved.

### Side-by-Side Mode (`--mode side-by-side`)

1. Create a new document
2. Insert a single-column table with two columns: Original | Translation
3. For each paragraph pair:
   - Left cell: deep copy the original paragraph's runs into the cell
   - Right cell: copy the original paragraph's runs, then replace text with translation — formatting matches the original paragraph
4. Copy page setup, margins, headers/footers from the original document

### Placeholder Protection

Legal documents often contain placeholders like `{{client_name}}` or `{{date}}`. The system prompt instructs the LLM to leave any `{{...}}` patterns untranslated. Additionally, a post-processing step re-scans the output and restores any `{{...}}` placeholders that were accidentally modified.

## Error Handling

| Scenario | Behavior |
|----------|----------|
| File not found | Print error, exit code 1 |
| Invalid DOCX | Print error, exit code 1 |
| LLM API error (auth) | Print error (check API key), exit code 1 |
| LLM API error (rate limit) | Retry with backoff, 3 attempts |
| LLM returns malformed output | Log warning, use sequential matching |
| Partial chunk failure | Log warning, produce partial output |
| Output path exists | Confirm overwrite or append timestamp suffix |

## Dependencies (v1)

```
python-docx
openai
python-dotenv
```

All installable via `pip install -r requirements.txt`.

## Out of Scope (v1)

- Database storage (SQLite) — v2
- Templating variables / structured data injection — v2
- React/TypeScript UI — v3
- FastAPI backend — v3
- Batch processing multiple files — v2
- Document history / audit trail — v2
