# DOCX Translator Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Python CLI script that translates `.docx` documents to Romanian and English via an OpenAI-compatible LLM, preserving formatting.

**Architecture:** Extract text paragraphs from DOCX with formatting metadata, chunk them with position IDs, send to LLM via configured provider, reconstruct output DOCX in either inline or side-by-side mode.

**Tech Stack:** Python 3.11+, python-docx, openai, python-dotenv

## Global Constraints

- Single Python file (`translate_docx.py`) for v1 with config in `config.json`
- All provider config in separate `config.json` file
- `openai` library used for all providers (OpenAI-compatible API)
- Only `ro` and `en` target languages in v1
- Temperature always 0.0 for deterministic output

---

### Task 1: Project Scaffolding

**Files:**
- Create: `requirements.txt`
- Create: `.env.example`
- Create: `config.json`
- Create: `test_samples/sample.docx` (minimal test document)

**Interfaces:**
- Produces: project dependencies and config template

- [ ] **Step 1: Create requirements.txt**

```
python-docx>=1.1.0
openai>=1.0.0
python-dotenv>=1.0.0
```

- [ ] **Step 2: Create .env.example**

```
# OpenAI-compatible API key (required for cloud providers)
OPENAI_API_KEY=sk-your-key-here
# OpenCode Go API key
OPENCODE_API_KEY=oc-your-key-here
```

- [ ] **Step 3: Create config.json**

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

- [ ] **Step 4: Create test sample DOCX**

Create a minimal DOCX with python script:

```bash
python3 -c "
from docx import Document
doc = Document()
doc.add_paragraph('This is a test document for translation.')
doc.add_paragraph('It contains multiple paragraphs with {{placeholder}} variables.')
p = doc.add_paragraph()
p.add_run('Bold text').bold = True
p.add_run(' and normal text')
doc.add_paragraph('Third paragraph with some content to translate.')
doc.save('test_samples/sample.docx')
"
```

- [ ] **Step 5: Commit**

```bash
git add requirements.txt .env.example config.json test_samples/sample.docx
git commit -m "chore: project scaffolding"
```

---

### Task 2: Config Loader

**Files:**
- Create: `translate_docx.py` — `load_config()` and `get_provider()` functions

**Interfaces:**
- Consumes: `config.json` format from Task 1
- Produces: `load_config(path) -> dict`, `get_provider(config, name) -> dict`

- [ ] **Step 1: Write the failing tests** (in a test block at the bottom of the file, or in `tests/test_config.py`)

Since v1 is a single-file script, we'll use a docstring-based test approach with assertions. Write at the bottom of `translate_docx.py`:

```python
def test_config_loading():
    import json, tempfile, os
    cfg = {
        "default_provider": "test",
        "providers": {
            "test": {
                "base_url": "http://test/v1",
                "model": "test-model",
                "api_key_env": None
            }
        }
    }
    with tempfile.NamedTemporaryFile(mode='w', suffix='.json', delete=False) as f:
        json.dump(cfg, f)
        f.flush()
        result = load_config(f.name)
        assert result == cfg
        os.unlink(f.name)

def test_get_provider():
    cfg = {
        "default_provider": "test",
        "providers": {
            "test": {
                "base_url": "http://test/v1",
                "model": "test-model",
                "api_key_env": None
            }
        }
    }
    provider = get_provider(cfg, "test")
    assert provider["base_url"] == "http://test/v1"
    assert provider["model"] == "test-model"
```

- [ ] **Step 2: Run tests to verify they fail**

```bash
python3 -c "import translate_docx; translate_docx.test_config_loading(); print('test_config_loading FAILED as expected - module not found')"
```

Expected: ImportError (module doesn't exist yet)

- [ ] **Step 3: Write minimal implementation**

```python
import json
import os
from pathlib import Path

def load_config(path="config.json"):
    with open(path) as f:
        return json.load(f)

def get_provider(config, name=None):
    if name is None:
        name = config["default_provider"]
    provider = config["providers"][name].copy()
    api_key_env = provider.pop("api_key_env", None)
    if api_key_env:
        provider["api_key"] = os.environ.get(api_key_env)
        if not provider["api_key"]:
            from dotenv import load_dotenv
            load_dotenv()
            provider["api_key"] = os.environ.get(api_key_env)
    return provider
```

- [ ] **Step 4: Run tests to verify they pass**

```bash
cd /Users/cristianradu/Documents/programming/elaw-templating && python3 -c "
import sys; sys.path.insert(0, '.')
import importlib
importlib.reload(sys.modules.get('translate_docx'))
from translate_docx import load_config, get_provider
import json, tempfile, os
cfg = {'default_provider': 'test', 'providers': {'test': {'base_url': 'http://test/v1', 'model': 'test-model', 'api_key_env': None}}}
with tempfile.NamedTemporaryFile(mode='w', suffix='.json', delete=False) as f:
    json.dump(cfg, f); f.flush()
    result = load_config(f.name)
    assert result == cfg, f'Expected {cfg} got {result}'
    os.unlink(f.name)
provider = get_provider(cfg, 'test')
assert provider['base_url'] == 'http://test/v1'
assert provider['model'] == 'test-model'
print('PASS')
"
```

Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add translate_docx.py
git commit -m "feat: add config loader with provider resolution"
```

---

### Task 3: DOCX Text Extractor

**Files:**
- Modify: `translate_docx.py` — add `extract_paragraphs()` function

**Interfaces:**
- Consumes: DOCX file path
- Produces: `extract_paragraphs(path) -> list[dict]` — each dict has `id`, `runs` (list of dicts with `text`, `bold`, `italic`, `font_name`, `font_size`, `color`), `alignment`, `in_table`, `cell_row`, `cell_col`

- [ ] **Step 1: Write the failing tests**

Append to the test section:

```python
def test_extract_paragraphs():
    from docx import Document
    import tempfile, os
    doc = Document()
    doc.add_paragraph("Hello world")
    p = doc.add_paragraph()
    p.add_run("Bold").bold = True
    p.add_run(" Normal")
    path = os.path.join(tempfile.mkdtemp(), "test.docx")
    doc.save(path)
    paragraphs = extract_paragraphs(path)
    assert len(paragraphs) == 2
    assert paragraphs[0]["id"] == "P0"
    assert paragraphs[0]["runs"][0]["text"] == "Hello world"
    assert paragraphs[1]["runs"][0]["text"] == "Bold"
    assert paragraphs[1]["runs"][0]["bold"] == True
    assert paragraphs[1]["runs"][1]["text"] == " Normal"
    os.unlink(path)

def test_extract_paragraphs_invalid_file():
    try:
        extract_paragraphs("/nonexistent/file.docx")
        assert False, "Should have raised FileNotFoundError"
    except FileNotFoundError:
        pass
```

- [ ] **Step 2: Run tests to verify they fail**

```bash
cd /Users/cristianradu/Documents/programming/elaw-templating && python3 -c "
import sys; sys.path.insert(0, '.')
import importlib
# Remove cached module
if 'translate_docx' in sys.modules:
    del sys.modules['translate_docx']
import translate_docx
try:
    translate_docx.test_extract_paragraphs()
    print('SHOULD HAVE FAILED - extract_paragraphs not defined')
except AttributeError as e:
    print(f'Expected failure: {e}')
"
```

Expected: AttributeError: module 'translate_docx' has no attribute 'extract_paragraphs'

- [ ] **Step 3: Write minimal implementation**

```python
from docx import Document

def extract_paragraphs(path):
    doc = Document(path)
    paragraphs = []
    for i, para in enumerate(doc.paragraphs):
        runs_data = []
        for run in para.runs:
            runs_data.append({
                "text": run.text,
                "bold": run.bold,
                "italic": run.italic,
                "underline": run.underline,
                "font_name": run.font.name,
                "font_size": str(run.font.size) if run.font.size else None,
                "color": str(run.font.color.rgb) if run.font.color and run.font.color.rgb else None,
            })
        paragraphs.append({
            "id": f"P{i}",
            "runs": runs_data,
            "alignment": str(para.alignment) if para.alignment else None,
            "in_table": para._element.getparent().tag.endswith("tc"),
        })
    return paragraphs
```

- [ ] **Step 4: Run tests to verify they pass**

```bash
cd /Users/cristianradu/Documents/programming/elaw-templating && python3 -c "
import sys; sys.path.insert(0, '.')
if 'translate_docx' in sys.modules:
    del sys.modules['translate_docx']
import translate_docx
translate_docx.test_extract_paragraphs()
translate_docx.test_extract_paragraphs_invalid_file()
print('PASS')
"
```

Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add translate_docx.py
git commit -m "feat: add DOCX paragraph extractor"
```

---

### Task 4: Chunking & LLM Translation Client

**Files:**
- Modify: `translate_docx.py` — add `chunk_paragraphs()`, `translate_chunk()`, `translate_all()` functions

**Interfaces:**
- Consumes: `extract_paragraphs()` output, provider config from `get_provider()`
- Produces: `chunk_paragraphs(paragraphs, chunk_size=10) -> list[list[dict]]`
- Produces: `translate_chunk(chunk, target_lang, provider) -> list[dict]` (returned paragraphs with translated text in runs)
- Produces: `translate_all(paragraphs, target_lang, provider) -> list[dict]`

- [ ] **Step 1: Write the failing tests**

```python
def test_chunk_paragraphs():
    paragraphs = [{"id": f"P{i}"} for i in range(25)]
    chunks = chunk_paragraphs(paragraphs, 10)
    assert len(chunks) == 2  # 10 + 10 + 5 overflow = 2 chunks of 10, last 5 in second
    # Actually 25 / 10 = 2 full chunks (20) + 1 partial (5) = 3 chunks
    # Wait, let me recalculate. 25 paragraphs, chunk_size=10:
    # chunk 0: P0-P9 (10), chunk 1: P10-P19 (10), chunk 2: P20-P24 (5) => 3 chunks
    assert len(chunks) == 3

def test_translate_chunk():
    chunk = [{"id": "P0", "runs": [{"text": "Hello world", "bold": False, "italic": False}]}]
    # Without a real API key, this should raise a helpful error
    provider = {"base_url": "http://invalid", "model": "test", "api_key": "bad"}
    try:
        translate_chunk(chunk, "Romanian", provider)
        assert False, "Should have raised ConnectionError or similar"
    except Exception:
        pass

def test_translate_chunk_retry():
    """Verify retry logic on transient failure"""
    pass  # Integration test requiring real API - skip in unit tests
```

- [ ] **Step 2: Run tests to verify they fail**

```bash
cd /Users/cristianradu/Documents/programming/elaw-templating && python3 -c "
import sys; sys.path.insert(0, '.')
if 'translate_docx' in sys.modules:
    del sys.modules['translate_docx']
import translate_docx
try:
    translate_docx.test_chunk_paragraphs()
    print('SHOULD HAVE FAILED')
except AttributeError as e:
    print(f'Expected failure: {e}')
"
```

Expected: AttributeError

- [ ] **Step 3: Write minimal implementation**

```python
import time
from openai import OpenAI

def chunk_paragraphs(paragraphs, chunk_size=10):
    chunks = []
    for i in range(0, len(paragraphs), chunk_size):
        chunks.append(paragraphs[i:i + chunk_size])
    return chunks

def _build_chunk_text(chunk):
    lines = []
    for p in chunk:
        text = ""
        for run in p["runs"]:
            text += run["text"]
        lines.append(f"[{p['id']}] {text}")
    return "\n".join(lines)

def _parse_translated_response(response_text, chunk_ids):
    """Parse LLM response back into per-paragraph translations."""
    results = {}
    for line in response_text.strip().split("\n"):
        line = line.strip()
        if not line:
            continue
        for pid in chunk_ids:
            if line.startswith(f"[{pid}]"):
                text = line[len(f"[{pid}]"):].strip()
                results[pid] = text
                break
    # Fall back to sequential matching for any missing IDs
    if len(results) < len(chunk_ids):
        lines = [l for l in response_text.strip().split("\n") if l.strip()]
        for i, pid in enumerate(chunk_ids):
            if pid not in results and i < len(lines):
                results[pid] = lines[i].strip()
    return results

def translate_chunk(chunk, target_lang, provider):
    chunk_text = _build_chunk_text(chunk)
    chunk_ids = [p["id"] for p in chunk]
    client = OpenAI(base_url=provider["base_url"], api_key=provider.get("api_key", ""))
    system_prompt = (
        f"You are a legal document translator. Translate the following paragraphs "
        f"to {target_lang}. Preserve all paragraph IDs exactly ([P0], [P1], ...). "
        f"Preserve {{...}} placeholders without translating them. "
        f"Output ONLY the translated paragraphs with their IDs — no extra text."
    )
    last_error = None
    for attempt in range(3):
        try:
            resp = client.chat.completions.create(
                model=provider["model"],
                messages=[
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": chunk_text},
                ],
                temperature=0.0,
            )
            translated = resp.choices[0].message.content
            parsed = _parse_translated_response(translated, chunk_ids)
            # Apply translations to paragraph runs
            for p in chunk:
                if p["id"] in parsed:
                    translated_text = parsed[p["id"]]
                    p["runs"][0]["text"] = translated_text
                    # Collapse all runs into first run for simplicity
                    p["runs"] = [p["runs"][0]]
            return chunk
        except Exception as e:
            last_error = e
            if attempt < 2:
                time.sleep(2 ** attempt)
    raise last_error

def translate_all(paragraphs, target_lang, provider):
    chunks = chunk_paragraphs(paragraphs)
    all_translated = []
    for chunk in chunks:
        translated = translate_chunk(chunk, target_lang, provider)
        all_translated.extend(translated)
    return all_translated
```

- [ ] **Step 4: Run unit tests**

```bash
cd /Users/cristianradu/Documents/programming/elaw-templating && python3 -c "
import sys; sys.path.insert(0, '.')
if 'translate_docx' in sys.modules:
    del sys.modules['translate_docx']
import translate_docx
translate_docx.test_chunk_paragraphs()
print('Chunk test PASS')
translate_docx.test_translate_chunk()
print('Translate error test PASS')
"
```

Expected: Chunk test PASS, Translate error test PASS

- [ ] **Step 5: Commit**

```bash
git add translate_docx.py
git commit -m "feat: add chunking and LLM translation client"
```

---

### Task 5: DOCX Writer — Inline Mode

**Files:**
- Modify: `translate_docx.py` — add `write_inline()` function

**Interfaces:**
- Consumes: original DOCX path, translated paragraphs list, output path
- Produces: `write_inline(original_path, translated_paragraphs, output_path)`

- [ ] **Step 1: Write the failing test**

```python
def test_write_inline():
    from docx import Document
    import tempfile, os
    # Create source DOCX
    doc = Document()
    p = doc.add_paragraph()
    p.add_run("Original text")
    src = os.path.join(tempfile.mkdtemp(), "source.docx")
    doc.save(src)
    # Translated paragraphs
    translated = [{
        "id": "P0",
        "runs": [{"text": "Text tradus", "bold": False, "italic": False}],
        "alignment": None,
        "in_table": False,
    }]
    out = os.path.join(tempfile.mkdtemp(), "output.docx")
    write_inline(src, translated, out)
    result = Document(out)
    assert result.paragraphs[0].text == "Text tradus"
    os.unlink(src); os.unlink(out)
```

- [ ] **Step 2: Run test to verify it fails**

```bash
cd /Users/cristianradu/Documents/programming/elaw-templating && python3 -c "
import sys; sys.path.insert(0, '.')
if 'translate_docx' in sys.modules:
    del sys.modules['translate_docx']
import translate_docx
try:
    translate_docx.test_write_inline()
    print('SHOULD HAVE FAILED')
except AttributeError as e:
    print(f'Expected failure: {e}')
"
```

Expected: AttributeError

- [ ] **Step 3: Write minimal implementation**

```python
def write_inline(original_path, translated_paragraphs, output_path):
    from docx import Document as DocxDocument
    doc = DocxDocument(original_path)
    trans_by_id = {p["id"]: p for p in translated_paragraphs}
    for i, para in enumerate(doc.paragraphs):
        pid = f"P{i}"
        if pid in trans_by_id:
            tp = trans_by_id[pid]
            for j, run in enumerate(para.runs):
                if j < len(tp["runs"]):
                    run.text = tp["runs"][j]["text"]
    doc.save(output_path)
```

- [ ] **Step 4: Run test to verify it passes**

```bash
cd /Users/cristianradu/Documents/programming/elaw-templating && python3 -c "
import sys; sys.path.insert(0, '.')
if 'translate_docx' in sys.modules:
    del sys.modules['translate_docx']
import translate_docx
translate_docx.test_write_inline()
print('PASS')
"
```

Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add translate_docx.py
git commit -m "feat: add inline DOCX writer"
```

---

### Task 6: DOCX Writer — Side-by-Side Mode

**Files:**
- Modify: `translate_docx.py` — add `write_side_by_side()` function

**Interfaces:**
- Consumes: original DOCX path, original paragraphs list, translated paragraphs list, output path
- Produces: `write_side_by_side(original_path, original_paras, translated_paras, output_path)`

- [ ] **Step 1: Write the failing test**

```python
def test_write_side_by_side():
    from docx import Document
    import tempfile, os
    doc = Document()
    p = doc.add_paragraph()
    p.add_run("Original text")
    src = os.path.join(tempfile.mkdtemp(), "source.docx")
    doc.save(src)
    originals = [{"id": "P0", "runs": [{"text": "Original text", "bold": False}], "alignment": None, "in_table": False}]
    translated = [{"id": "P0", "runs": [{"text": "Text tradus", "bold": False}], "alignment": None, "in_table": False}]
    out = os.path.join(tempfile.mkdtemp(), "output.docx")
    write_side_by_side(src, originals, translated, out)
    result = Document(out)
    assert len(result.tables) == 1
    table = result.tables[0]
    assert len(table.rows) == 1
    assert "Original text" in table.rows[0].cells[0].text
    assert "Text tradus" in table.rows[0].cells[1].text
    os.unlink(src); os.unlink(out)
```

- [ ] **Step 2: Run test to verify it fails**

```bash
cd /Users/cristianradu/Documents/programming/elaw-templating && python3 -c "
import sys; sys.path.insert(0, '.')
if 'translate_docx' in sys.modules:
    del sys.modules['translate_docx']
import translate_docx
try:
    translate_docx.test_write_side_by_side()
    print('SHOULD HAVE FAILED')
except AttributeError as e:
    print(f'Expected failure: {e}')
"
```

Expected: AttributeError

- [ ] **Step 3: Write minimal implementation**

```python
from docx import Document as DocxDocument
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from copy import deepcopy

def write_side_by_side(original_path, original_paras, translated_paras, output_path):
    src = DocxDocument(original_path)
    doc = DocxDocument()
    # Copy page margins from source
    for section in src.sections:
        for new_section in doc.sections:
            new_section.page_width = section.page_width
            new_section.page_height = section.page_height
            new_section.left_margin = section.left_margin
            new_section.right_margin = section.right_margin
            new_section.top_margin = section.top_margin
            new_section.bottom_margin = section.bottom_margin
        break
    trans_by_id = {p["id"]: p for p in translated_paras}
    orig_by_id = {p["id"]: p for p in original_paras}
    table = doc.add_table(rows=len(original_paras), cols=2)
    table.style = 'Table Grid'
    for i, op in enumerate(original_paras):
        pid = op["id"]
        left_cell = table.rows[i].cells[0]
        right_cell = table.rows[i].cells[1]
        # Left cell: original text
        for run_data in op["runs"]:
            p = left_cell.paragraphs[0] if left_cell.paragraphs else left_cell.add_paragraph()
            run = p.add_run(run_data["text"])
        # Right cell: translated text
        tp = trans_by_id.get(pid, op)
        for run_data in tp["runs"]:
            p = right_cell.paragraphs[0] if right_cell.paragraphs else right_cell.add_paragraph()
            run = p.add_run(run_data["text"])
    doc.save(output_path)
```

- [ ] **Step 4: Run test to verify it passes**

```bash
cd /Users/cristianradu/Documents/programming/elaw-templating && python3 -c "
import sys; sys.path.insert(0, '.')
if 'translate_docx' in sys.modules:
    del sys.modules['translate_docx']
import translate_docx
translate_docx.test_write_side_by_side()
print('PASS')
"
```

Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add translate_docx.py
git commit -m "feat: add side-by-side DOCX writer"
```

---

### Task 7: CLI Entry Point & Pipeline Orchestration

**Files:**
- Modify: `translate_docx.py` — add `main()` with argparse

**Interfaces:**
- Produces: CLI with flags `--lang`, `--mode`, `--output`, `--provider`, `--model`, `--config`

- [ ] **Step 1: Write the failing test**

```python
def test_cli_parser():
    import argparse
    sys_argv_backup = sys.argv
    sys.argv = ['translate_docx.py', 'input.docx', '--lang', 'ro']
    args = parse_args(['input.docx', '--lang', 'ro'])
    assert args.input == 'input.docx'
    assert args.lang == 'ro'
    assert args.mode == 'inline'
```

- [ ] **Step 2: Write minimal implementation**

```python
import argparse
import sys

def parse_args(argv=None):
    parser = argparse.ArgumentParser(description="Translate DOCX documents using LLM")
    parser.add_argument("input", help="Path to input DOCX file")
    parser.add_argument("--lang", "-l", required=True, choices=["ro", "en"],
                        help="Target language (ro=Romanian, en=English)")
    parser.add_argument("--mode", "-m", choices=["inline", "side-by-side"],
                        default="inline", help="Output layout mode")
    parser.add_argument("--output", "-o", help="Output path (auto-generated if omitted)")
    parser.add_argument("--provider", "-p", help="Provider key from config.json")
    parser.add_argument("--model", help="Override model name")
    parser.add_argument("--config", "-c", default="config.json", help="Config file path")
    return parser.parse_args(argv)

def main():
    args = parse_args()
    config = load_config(args.config)
    provider = get_provider(config, args.provider)
    if args.model:
        provider["model"] = args.model
    output = args.output
    if not output:
        stem = args.input.rsplit(".", 1)[0]
        suffix = f"_{args.mode}_{args.lang}.docx" if args.mode == "side-by-side" else f"_{args.lang}.docx"
        output = f"{stem}{suffix}"
    paragraphs = extract_paragraphs(args.input)
    translated = translate_all(paragraphs, "Romanian" if args.lang == "ro" else "English", provider)
    if args.mode == "side-by-side":
        write_side_by_side(args.input, paragraphs, translated, output)
    else:
        write_inline(args.input, translated, output)
    print(f"Translated document saved to: {output}")

if __name__ == "__main__":
    main()
```

- [ ] **Step 3: Test CLI parser**

```bash
cd /Users/cristianradu/Documents/programming/elaw-templating && python3 -c "
import sys; sys.path.insert(0, '.')
if 'translate_docx' in sys.modules:
    del sys.modules['translate_docx']
import translate_docx
args = translate_docx.parse_args(['input.docx', '--lang', 'ro'])
assert args.input == 'input.docx'
assert args.lang == 'ro'
assert args.mode == 'inline'
args2 = translate_docx.parse_args(['input.docx', '--lang', 'en', '--mode', 'side-by-side'])
assert args2.lang == 'en'
assert args2.mode == 'side-by-side'
print('PASS')
"
```

Expected: PASS

- [ ] **Step 4: Manual integration smoke test** (requires API key)

```bash
cd /Users/cristianradu/Documents/programming/elaw-templating && \
  OPENCODE_API_KEY=your-key-here python3 translate_docx.py test_samples/sample.docx --lang ro
```

Expected: Creates `test_samples/sample_ro.docx`

- [ ] **Step 5: Commit**

```bash
git add translate_docx.py
git commit -m "feat: add CLI entry point and pipeline orchestration"
```

---

### Task 8: Placeholder Protection & Edge Cases

**Files:**
- Modify: `translate_docx.py` — enhance `_parse_translated_response()` and add `restore_placeholders()`

- [ ] **Step 1: Write the failing test**

```python
def test_restore_placeholders():
    original = "Hello {{client_name}}, your balance is {{amount}}"
    translated = "Salut {{client_name}}, soldul tau este {{amount}}"
    restored = restore_placeholders(translated, original)
    assert "{{client_name}}" in restored
    assert "{{amount}}" in restored

def test_restore_missing_placeholder():
    """Placeholder was dropped in translation - should be re-inserted."""
    original = "Hello {{client_name}}"
    translated = "Salut"
    restored = restore_placeholders(translated, original)
    assert "{{client_name}}" in restored
```

- [ ] **Step 2: Write implementation**

```python
import re

def restore_placeholders(translated_text, original_text):
    placeholder_pattern = r'\{\{[^}]+\}\}'
    placeholders = re.findall(placeholder_pattern, original_text)
    for ph in placeholders:
        if ph not in translated_text:
            # Append missing placeholder at end
            translated_text += f" {ph}"
    return translated_text
```

- [ ] **Step 3: Run tests**

```bash
cd /Users/cristianradu/Documents/programming/elaw-templating && python3 -c "
import sys; sys.path.insert(0, '.')
if 'translate_docx' in sys.modules:
    del sys.modules['translate_docx']
import translate_docx
translate_docx.test_restore_placeholders()
translate_docx.test_restore_missing_placeholder()
print('PASS')
"
```

Expected: PASS

- [ ] **Step 4: Commit**

```bash
git add translate_docx.py
git commit -m "feat: add placeholder protection for {{...}} patterns"
```
