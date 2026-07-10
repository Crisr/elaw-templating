# Transform2Cell Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add `--transform2cell` CLI flag that converts Word 2-column layout documents to table-based 2-column documents, with frontend detection and UI adaptation.

**Architecture:** Three-tier fallback: exact column-break marker → heuristic split + AI verify → full AI matching. Frontend uses JSZip to detect columns client-side.

**Tech Stack:** Python 3.11+, python-docx, openai, React 18, JSZip 3

## Global Constraints

- All new functions added to existing `translate_docx.py`
- Test functions follow the existing pattern (embedded in same file, run when invoked without args)
- Tests must not require network access (mock AI calls)
- Frontend detects 2-column layout on file drop using JSZip (no backend call needed)

---

### Task 1: Test Fixture + Core Detection Functions

**Files:**
- Modify: `translate_docx.py` — add test fixture generator, detection functions, and their tests

**Interfaces:**
- Produces: `_create_2column_test_docx(path)` — creates a synthetic DOCX with 2-column layout and column break marker
- Produces: `_has_2column_layout(doc) -> bool` — checks section `<w:cols w:num="2">`
- Produces: `_find_column_break(doc) -> int | None` — finds `<w:br w:type="column"/>` index
- Produces: `_heuristic_column_split(doc) -> int` — returns `len(paras) // 2`
- Produces: `_pair_by_position(paras, split_idx) -> list[tuple]` — splits and zips by index

- [ ] **Step 1: Write test fixture generator**

At the top of `translate_docx.py` (after existing imports), no changes needed. Add this new utility function:

```python
def _create_2column_test_docx(path):
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn

    doc = Document()
    sect_pr = doc.sections[0]._sectPr
    cols = OxmlElement('w:cols')
    cols.set(qn('w:num'), '2')
    cols.set(qn('w:space'), '720')
    sect_pr.append(cols)

    doc.add_paragraph("First original paragraph.")
    doc.add_paragraph("Second original paragraph.")
    doc.add_paragraph("Third original paragraph.")

    p_break = doc.add_paragraph()
    r_break = p_break.add_run()
    br = OxmlElement('w:br')
    br.set(qn('w:type'), 'column')
    r_break._element.append(br)

    doc.add_paragraph("Primul paragraf original.")
    doc.add_paragraph("Al doilea paragraf original.")
    doc.add_paragraph("Al treilea paragraf original.")
    doc.save(path)
```

- [ ] **Step 2: Write tests for detection functions**

Add to the test block at the bottom of `translate_docx.py`:

```python
def test_has_2column_layout():
    import tempfile, os
    from docx import Document
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    doc = Document()
    path = os.path.join(tempfile.mkdtemp(), "test.docx")
    # Without 2-column layout
    doc.save(path)
    assert not _has_2column_layout(Document(path))
    # With 2-column layout
    doc2 = Document()
    sect_pr = doc2.sections[0]._sectPr
    cols = OxmlElement('w:cols')
    cols.set(qn('w:num'), '2')
    cols.set(qn('w:space'), '720')
    sect_pr.append(cols)
    path2 = os.path.join(tempfile.mkdtemp(), "test2.docx")
    doc2.save(path2)
    assert _has_2column_layout(Document(path2))
    os.unlink(path); os.unlink(path2)

def test_find_column_break():
    import tempfile, os
    path = os.path.join(tempfile.mkdtemp(), "test.docx")
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    doc = Document()
    doc.add_paragraph("Before break")
    p = doc.add_paragraph()
    r = p.add_run()
    br = OxmlElement('w:br')
    br.set(qn('w:type'), 'column')
    r._element.append(br)
    doc.add_paragraph("After break")
    doc.save(path)
    idx = _find_column_break(Document(path))
    assert idx == 1, f"Expected 1, got {idx}"
    # No break marker
    doc2 = Document()
    doc2.add_paragraph("No break")
    doc2.add_paragraph("No break either")
    path2 = os.path.join(tempfile.mkdtemp(), "test2.docx")
    doc2.save(path2)
    assert _find_column_break(Document(path2)) is None
    os.unlink(path); os.unlink(path2)

def test_heuristic_column_split():
    from docx import Document
    doc = Document()
    for _ in range(10):
        doc.add_paragraph("x")
    assert _heuristic_column_split(doc) == 5
    doc2 = Document()
    for _ in range(7):
        doc2.add_paragraph("x")
    assert _heuristic_column_split(doc2) == 3

def test_pair_by_position():
    paras = [f"P{i}" for i in range(6)]
    pairs = _pair_by_position(paras, 3)
    assert len(pairs) == 3
    assert pairs[0] == ("P0", "P3")
    assert pairs[1] == ("P1", "P4")
    assert pairs[2] == ("P2", "P5")
    # Uneven
    paras2 = [f"P{i}" for i in range(5)]
    pairs2 = _pair_by_position(paras2, 2)
    assert len(pairs2) == 3
    assert pairs2[0] == ("P0", "P2")
    assert pairs2[1] == ("P1", "P3")
    assert pairs2[2] == (None, "P4")
```

- [ ] **Step 3: Run tests to verify they fail**

```bash
cd /Users/cristianradu/Documents/programming/elaw-templating && python3 -c "
import sys; sys.path.insert(0, '.')
if 'translate_docx' in sys.modules:
    del sys.modules['translate_docx']
import translate_docx
for name in ['test_has_2column_layout', 'test_find_column_break', 'test_heuristic_column_split', 'test_pair_by_position']:
    try:
        getattr(translate_docx, name)()
        print(f'{name}: SHOULD HAVE FAILED')
    except AttributeError as e:
        print(f'{name}: Expected failure: {e}')
"
```

Expected: 4 AttributeErrors

- [ ] **Step 4: Implement detection functions**

Add before the test block:

```python
def _has_2column_layout(doc):
    for section in doc.sections:
        sect_pr = section._sectPr
        if sect_pr is not None:
            cols = sect_pr.find(f'{NS_W}cols')
            if cols is not None:
                num = cols.get(f'{NS_W}num')
                if num and int(num) == 2:
                    return True
    return False


def _find_column_break(doc):
    for i, para in enumerate(doc.paragraphs):
        for br in para._element.findall(f'.//{NS_W}br'):
            if br.get(f'{NS_W}type') == 'column':
                return i
    return None


def _heuristic_column_split(doc):
    return len(doc.paragraphs) // 2


def _pair_by_position(paras, split_idx):
    col1 = list(paras[:split_idx])
    col2 = list(paras[split_idx:])
    pairs = []
    for i in range(max(len(col1), len(col2))):
        p1 = col1[i] if i < len(col1) else None
        p2 = col2[i] if i < len(col2) else None
        pairs.append((p1, p2))
    return pairs
```

- [ ] **Step 5: Run tests to verify they pass**

```bash
cd /Users/cristianradu/Documents/programming/elaw-templating && python3 -c "
import sys; sys.path.insert(0, '.')
if 'translate_docx' in sys.modules:
    del sys.modules['translate_docx']
import translate_docx
translate_docx.test_has_2column_layout()
print('test_has_2column_layout PASS')
translate_docx.test_find_column_break()
print('test_find_column_break PASS')
translate_docx.test_heuristic_column_split()
print('test_heuristic_column_split PASS')
translate_docx.test_pair_by_position()
print('test_pair_by_position PASS')
"
```

Expected: 4 PASS

- [ ] **Step 6: Commit**

```bash
git add translate_docx.py
git commit -m "feat: add column detection and split functions for transform2cell"
```

---

### Task 2: Table Writer

**Files:**
- Modify: `translate_docx.py` — add `_write_2cell_table()` function

**Interfaces:**
- Consumes: `pairs: list[tuple[dict | None, dict | None]]` — (left_data, right_data) pairs
- Produces: `_write_2cell_table(original_path, pairs, output_path)` — saves borderless 2-col table DOCX

- [ ] **Step 1: Write the failing test**

```python
def test_write_2cell_table():
    import tempfile, os
    from docx import Document
    doc = Document()
    doc.add_paragraph("Source only")
    src = os.path.join(tempfile.mkdtemp(), "src.docx")
    doc.save(src)
    pairs = [
        ({"runs": [{"text": "Left A", "bold": False}], "alignment": None, "numpr": None},
         {"runs": [{"text": "Right A", "bold": False}], "alignment": None, "numpr": None}),
        ({"runs": [{"text": "Left B", "bold": False}], "alignment": None, "numpr": None},
         {"runs": [{"text": "Right B", "bold": False}], "alignment": None, "numpr": None}),
    ]
    out = os.path.join(tempfile.mkdtemp(), "out.docx")
    _write_2cell_table(src, pairs, out)
    result = Document(out)
    assert len(result.tables) == 1
    t = result.tables[0]
    assert len(t.rows) == 2
    assert t.rows[0].cells[0].text == "Left A"
    assert t.rows[0].cells[1].text == "Right A"
    assert t.rows[1].cells[0].text == "Left B"
    assert t.rows[1].cells[1].text == "Right B"
    os.unlink(src); os.unlink(out)

def test_write_2cell_table_no_borders():
    import tempfile, os
    from docx import Document
    from lxml import etree
    doc = Document()
    doc.add_paragraph("x")
    src = os.path.join(tempfile.mkdtemp(), "src.docx")
    doc.save(src)
    pairs = [({"runs": [{"text": "A", "bold": False}], "alignment": None, "numpr": None},
              {"runs": [{"text": "B", "bold": False}], "alignment": None, "numpr": None})]
    out = os.path.join(tempfile.mkdtemp(), "out.docx")
    _write_2cell_table(src, pairs, out)
    result = Document(out)
    tbl = result.tables[0]
    tbl_pr = tbl._tbl.tblPr
    borders = tbl_pr.find(f'{NS_W}tblBorders')
    assert borders is not None, "Table should have tblBorders element"
    os.unlink(src); os.unlink(out)
```

- [ ] **Step 2: Run tests to verify they fail**

```bash
cd /Users/cristianradu/Documents/programming/elaw-templating && python3 -c "
import sys; sys.path.insert(0, '.')
if 'translate_docx' in sys.modules:
    del sys.modules['translate_docx']
import translate_docx
try:
    translate_docx.test_write_2cell_table()
    print('SHOULD HAVE FAILED')
except AttributeError as e:
    print(f'Expected: {e}')
"
```

Expected: AttributeError

- [ ] **Step 3: Implement `_write_2cell_table`**

```python
def _write_2cell_table(original_path, pairs, output_path):
    src = Document(original_path)
    doc = Document()
    for section in src.sections:
        for new_section in doc.sections:
            new_section.page_width = section.page_width
            new_section.page_height = section.page_height
            new_section.left_margin = section.left_margin
            new_section.right_margin = section.right_margin
            new_section.top_margin = section.top_margin
            new_section.bottom_margin = section.bottom_margin
        break

    table = doc.add_table(rows=len(pairs), cols=2)
    ns = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
    tbl_pr = table._tbl.tblPr
    tbl_borders = doc.element.makeelement(ns + 'tblBorders', {})
    for edge in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
        child = doc.element.makeelement(ns + edge, {})
        child.set(ns + 'val', 'none')
        child.set(ns + 'sz', '0')
        child.set(ns + 'space', '0')
        child.set(ns + 'color', 'auto')
        tbl_borders.append(child)
    tbl_pr.append(tbl_borders)

    for i, (left_data, right_data) in enumerate(pairs):
        left_cell = table.rows[i].cells[0]
        right_cell = table.rows[i].cells[1]
        if left_data:
            for run_data in left_data["runs"]:
                p = left_cell.paragraphs[0] if left_cell.paragraphs else left_cell.add_paragraph()
                run = p.add_run(run_data["text"])
                _apply_run_formatting(run, run_data)
        if right_data:
            for run_data in right_data["runs"]:
                p = right_cell.paragraphs[0] if right_cell.paragraphs else right_cell.add_paragraph()
                run = p.add_run(run_data["text"])
                _apply_run_formatting(run, run_data)
    doc.save(output_path)
```

- [ ] **Step 4: Run tests to verify they pass**

```bash
cd /Users/cristianradu/Documents/programming/elaw-templating && python3 -c "
import sys; sys.path.insert(0, '.')
if 'translate_docx' in sys.modules:
    del sys.modules['translate_docx']
import translate_docx
translate_docx.test_write_2cell_table()
print('test_write_2cell_table PASS')
translate_docx.test_write_2cell_table_no_borders()
print('test_write_2cell_table_no_borders PASS')
"
```

Expected: 2 PASS

- [ ] **Step 5: Commit**

```bash
git add translate_docx.py
git commit -m "feat: add 2-cell table writer for transform2cell"
```

---

### Task 3: AI Verification + Orchestrator + CLI

**Files:**
- Modify: `translate_docx.py` — add `_llm_verify_pairs()`, `_llm_full_column_matching()`, `transform2cell()`, update `parse_args()`, `main()`, add messages

**Interfaces:**
- Produces: `_llm_verify_pairs(col1_texts, col2_texts, pairs, provider) -> list[tuple]`
- Produces: `_llm_full_column_matching(all_texts, provider) -> list[tuple]`
- Produces: `transform2cell(input_path, output_path, provider=None) -> None`
- Updates: `parse_args(argv)` — adds `--transform2cell` flag
- Updates: `main()` — dispatches to `transform2cell` when flag set

- [ ] **Step 1: Write the failing tests**

```python
def test_transform2cell_integration():
    import tempfile, os
    from docx import Document
    path = os.path.join(tempfile.mkdtemp(), "src.docx")
    _create_2column_test_docx(path)
    out = os.path.join(tempfile.mkdtemp(), "out.docx")
    transform2cell(path, out)
    result = Document(out)
    assert len(result.tables) == 1
    t = result.tables[0]
    assert len(t.rows) == 3
    assert "First original" in t.rows[0].cells[0].text
    assert "Primul paragraf" in t.rows[0].cells[1].text
    assert "Second original" in t.rows[1].cells[0].text
    assert "Al doilea" in t.rows[1].cells[1].text
    assert "Third original" in t.rows[2].cells[0].text
    assert "Al treilea" in t.rows[2].cells[1].text
    os.unlink(path); os.unlink(out)

def test_cli_parser_transform2cell():
    args = parse_args(['input.docx', '--transform2cell'])
    assert args.transform2cell is True
    assert args.input == 'input.docx'
    # Should not require --lang
    args2 = parse_args(['input.docx', '--transform2cell', '--output', 'out.docx'])
    assert args2.transform2cell
    assert args2.output == 'out.docx'
```

- [ ] **Step 2: Run tests to verify they fail**

```bash
cd /Users/cristianradu/Documents/programming/elaw-templating && python3 -c "
import sys; sys.path.insert(0, '.')
if 'translate_docx' in sys.modules:
    del sys.modules['translate_docx']
import translate_docx
try:
    translate_docx.test_transform2cell_integration()
    print('SHOULD HAVE FAILED')
except AttributeError as e:
    print(f'Expected: {e}')
try:
    translate_docx.test_cli_parser_transform2cell()
    print('SHOULD HAVE FAILED')
except (AttributeError, SystemExit) as e:
    print(f'Expected: {e}')
"
```

Expected: AttributeError for transform2cell, maybe SystemExit for CLI (if argparse fails)

- [ ] **Step 3: Implement `_llm_verify_pairs`**

```python
def _llm_verify_pairs(col1_dicts, col2_dicts, pairs, provider):
    if provider is None:
        return pairs
    lines = []
    for i, (p1, p2) in enumerate(pairs):
        t1 = "".join(r["text"] for r in p1["runs"]) if p1 else ""
        t2 = "".join(r["text"] for r in p2["runs"]) if p2 else ""
        lines.append(f"[ROW {i}] Left: {t1}")
        lines.append(f"[ROW {i}] Right: {t2}")
    chunk_text = "\n".join(lines)
    client = OpenAI(base_url=provider["base_url"], api_key=provider.get("api_key", "not-needed"))
    system_prompt = (
        "You are verifying a two-column document alignment. "
        "Column 1 (Left) has original text, Column 2 (Right) has the translation. "
        "Each ROW shows a paired paragraph. If any rows are misaligned "
        "(wrong original matched to translation), return corrected pairs.\n\n"
        'Return ONLY JSON: {"pairs": [[0,0], [1,1], ...], '
        '"confidence": "high|medium|low", "issues": ["note any mismatches"]}'
    )
    try:
        resp = client.chat.completions.create(
            model=provider["model"],
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": chunk_text},
            ],
            temperature=0.0,
        )
        result = resp.choices[0].message.content
        parsed = json.loads(result)
        confidence = parsed.get("confidence", "low")
        if confidence == "low":
            print("Warning: AI verification confidence low, using heuristic pairing", file=sys.stderr)
        return pairs
    except Exception as e:
        print(f"Warning: AI verification failed ({e}), using heuristic pairing", file=sys.stderr)
        return pairs
```

- [ ] **Step 4: Implement `_llm_full_column_matching`**

```python
def _llm_full_column_matching(all_dicts, provider):
    if provider is None:
        return None
    lines = []
    for i, p in enumerate(all_dicts):
        text = "".join(r["text"] for r in p["runs"])
        lines.append(f"[P{i}] {text}")
    chunk_text = "\n".join(lines)
    client = OpenAI(base_url=provider["base_url"], api_key=provider.get("api_key", "not-needed"))
    system_prompt = (
        "This document has two columns. Paragraphs are listed in order "
        "(first column 1, then column 2). Determine the column split "
        "and pair each original with its translation.\n\n"
        'Return ONLY JSON: {"pairs": [[0, 5], [1, 6], ...], '
        '"confidence": "high|medium|low"} where each pair maps a column-1 '
        "paragraph index to its column-2 paragraph index."
    )
    try:
        resp = client.chat.completions.create(
            model=provider["model"],
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": chunk_text},
            ],
            temperature=0.0,
        )
        result = resp.choices[0].message.content
        parsed = json.loads(result)
        index_pairs = parsed.get("pairs", [])
        result_pairs = []
        for c1_idx, c2_idx in index_pairs:
            left = all_dicts[c1_idx] if c1_idx < len(all_dicts) else None
            right = all_dicts[c2_idx] if c2_idx < len(all_dicts) else None
            result_pairs.append((left, right))
        return result_pairs
    except Exception as e:
        print(f"Warning: full AI column matching failed ({e})", file=sys.stderr)
        return None
```

- [ ] **Step 5: Implement `transform2cell` orchestrator**

```python
def transform2cell(input_path, output_path, provider=None):
    doc = Document(input_path)
    if not _has_2column_layout(doc):
        print("Warning: Input document does not appear to have a 2-column layout. Results may be unexpected.", file=sys.stderr)

    split_idx = _find_column_break(doc)
    if split_idx is None:
        split_idx = _heuristic_column_split(doc)
        print(f"Info: No column break marker found. Using heuristic split at paragraph {split_idx}.", file=sys.stderr)

    all_dicts = [_build_para_dict(p) for p in doc.paragraphs]

    col1 = all_dicts[:split_idx]
    col2 = all_dicts[split_idx:]
    pairs = list(zip(col1, col2))

    if len(col1) != len(col2):
        print(f"Warning: Column paragraph counts differ (col1={len(col1)}, col2={len(col2)}).", file=sys.stderr)

    if provider:
        pairs = _llm_verify_pairs(col1, col2, pairs, provider)
        if pairs is None:
            pairs = _llm_full_column_matching(all_dicts, provider)
        if pairs is None:
            print("AI matching failed. Falling back to heuristic pairing.", file=sys.stderr)
            pairs = list(zip(col1, col2))

    _write_2cell_table(input_path, pairs, output_path)
```

- [ ] **Step 6: Add `_build_para_dict` helper**

```python
def _build_para_dict(para):
    return {
        "runs": _extract_runs(para),
        "alignment": str(para.alignment) if para.alignment else None,
        "numpr": _extract_numpr(para),
    }
```

- [ ] **Step 7: Update `parse_args` to add `--transform2cell`**

In `parse_args()`:

```python
parser.add_argument("--transform2cell", action="store_true",
                    help="Convert 2-column Word layout to table-based 2-column document")
```

Make `--lang` not required when `--transform2cell` is set by adding check in `main()` instead:

Replace:
```python
parser.add_argument("--lang", "-l", required=True, choices=["ro", "en"],
                    help=_["cli_lang_help"])
```

With:
```python
parser.add_argument("--lang", "-l", choices=["ro", "en"],
                    help=_["cli_lang_help"])
```

- [ ] **Step 8: Update `main()` to handle transform2cell**

```python
def main():
    args = parse_args()
    config = load_config(args.config)

    if args.transform2cell:
        provider = get_provider(config, args.provider) if args.provider else None
        if args.model and provider:
            provider["model"] = args.model
        output = args.output
        if not output:
            stem = args.input.rsplit(".", 1)[0]
            output = f"{stem}_2cell.docx"
        if os.path.exists(output):
            stem = output.rsplit(".docx", 1)[0]
            timestamp = time.strftime("%Y%m%d_%H%M%S")
            output = f"{stem}_{timestamp}.docx"
        transform2cell(args.input, output, provider)
        print(_["saved"].format(output=output))
        return

    if not args.lang:
        parser.error("--lang is required (unless using --transform2cell)")
    # ... rest of existing main()
```

- [ ] **Step 9: Add messages**

In `messages.py`:

```python
"warn_no_2column_layout": "Warning: Input document does not appear to have a 2-column layout. Results may be unexpected.",
"info_heuristic_column_split": "Info: No column break marker found. Using heuristic split at paragraph {idx}.",
"warn_transform2cell_mismatch": "Warning: Column paragraph counts differ (col1={c1}, col2={c2}).",
```

Update `"cli_mode_help"` and add:
```python
"cli_transform2cell_help": "Convert 2-column Word layout to table-based document",
```

- [ ] **Step 10: Run integration test**

```bash
cd /Users/cristianradu/Documents/programming/elaw-templating && python3 -c "
import sys; sys.path.insert(0, '.')
if 'translate_docx' in sys.modules:
    del sys.modules['translate_docx']
import translate_docx
translate_docx.test_transform2cell_integration()
print('test_transform2cell_integration PASS')
translate_docx.test_cli_parser_transform2cell()
print('test_cli_parser_transform2cell PASS')
"
```

Expected: 2 PASS

- [ ] **Step 11: Manual smoke test**

```bash
cd /Users/cristianradu/Documents/programming/elaw-templating && python3 -c "
import sys; sys.path.insert(0, '.')
if 'translate_docx' in sys.modules:
    del sys.modules['translate_docx']
import translate_docx
import tempfile, os
path = os.path.join(tempfile.mkdtemp(), 'test.docx')
translate_docx._create_2column_test_docx(path)
out = path.replace('.docx', '_2cell.docx')
translate_docx.transform2cell(path, out)
from docx import Document
d = Document(out)
print(f'Output has {len(d.tables)} table(s), {len(d.tables[0].rows)} row(s)')
for r in d.tables[0].rows:
    print(f'  Left: {r.cells[0].text[:50]} | Right: {r.cells[1].text[:50]}')
os.unlink(path); os.unlink(out)
"
```

Expected: 1 table, 3 rows, each row showing original | translation pairs

- [ ] **Step 12: Commit**

```bash
git add translate_docx.py messages.py
git commit -m "feat: add transform2cell orchestrator with AI verification and CLI flag"
```

---

### Task 4: Backend Integration

**Files:**
- Modify: `server.py` — accept `transform2cell` form field
- Modify: `worker.py` — handle `mode == "transform2cell"`

- [ ] **Step 1: Update server.py**

Add `transform2cell: bool = Form(False)` parameter to `api_translate`. Skip `lang`/`mode` validation when `transform2cell=True`. Store mode as `"transform2cell"`:

```python
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
    # ... rest unchanged
```

- [ ] **Step 2: Update worker.py**

In `_process()`, add handling for `"transform2cell"` mode:

```python
if mode == "transform2cell":
    p = translate_docx.get_provider(config, provider_name) if provider_name else None
    if model_override and p:
        p["model"] = model_override
    db.update_job(job_id, status="running", progress=0)
    result_path = db.UPLOAD_DIR / f"{job_id}_result.docx"
    translate_docx.transform2cell(source_path, str(result_path), p)
    db.update_job(job_id, status="done", progress=100, result_file=str(result_path))
    db.enforce_file_limit()
    return
```

Insert this block right after `db.update_job(job_id, status="running", progress=0)` and before the existing `paragraphs = translate_docx.extract_paragraphs(source_path)` line.

Also fix the download filename in server.py for transform2cell:

In `api_download()`:
```python
filename = f"transformed_cell.docx" if job["mode"] == "transform2cell" else f"translated_{job['language']}.docx"
```

- [ ] **Step 3: Verify server starts**

```bash
cd /Users/cristianradu/Documents/programming/elaw-templating && python3 -c "
import sys; sys.path.insert(0, '.')
from server import app
print('Server imported successfully')
"
```

Expected: "Server imported successfully"

- [ ] **Step 4: Commit**

```bash
git add server.py worker.py
git commit -m "feat: add transform2cell support to backend API and worker"
```

---

### Task 5: Frontend — JSZip + Column Detection

**Files:**
- Modify: `frontend/package.json` — add `jszip` dependency
- Modify: `frontend/src/components/DropZone.tsx` — add 2-column detection

- [ ] **Step 1: Add JSZip dependency**

```bash
cd /Users/cristianradu/Documents/programming/elaw-templating/frontend && npm install jszip
```

- [ ] **Step 2: Update DropZone.tsx**

Add `onTwoColumnDetected` prop and JSZip detection logic:

```tsx
import { useCallback, useRef, useState } from 'react'
import { useLocale } from '../LocaleContext'
import JSZip from 'jszip'

interface Props {
  file: File | null
  onFile: (f: File) => void
  disabled: boolean
  onTwoColumnDetected?: (v: boolean) => void
}

export default function DropZone({ file, onFile, disabled, onTwoColumnDetected }: Props) {
  const [dragging, setDragging] = useState(false)
  const inputRef = useRef<HTMLInputElement>(null)
  const { messages } = useLocale()
  const msg = messages.dropZone

  const detectTwoColumn = useCallback(async (f: File) => {
    if (!onTwoColumnDetected) return
    try {
      const buf = await f.arrayBuffer()
      const zip = await JSZip.loadAsync(buf)
      const docXml = await zip.file('word/document.xml')?.async('string')
      if (!docXml) { onTwoColumnDetected(false); return }
      const has2Cols = /<w:cols[^>]*w:num\s*=\s*"2"/.test(docXml)
      onTwoColumnDetected(has2Cols)
    } catch {
      onTwoColumnDetected(false)
    }
  }, [onTwoColumnDetected])

  const handleDrop = useCallback(
    (e: React.DragEvent) => {
      e.preventDefault()
      setDragging(false)
      if (disabled) return
      const f = e.dataTransfer.files[0]
      if (f && f.name.endsWith('.docx')) {
        onFile(f)
        detectTwoColumn(f)
      }
    },
    [disabled, onFile, detectTwoColumn]
  )

  const handleChange = useCallback(
    (e: React.ChangeEvent<HTMLInputElement>) => {
      const f = e.target.files?.[0]
      if (f) {
        onFile(f)
        detectTwoColumn(f)
      }
    },
    [onFile, detectTwoColumn]
  )

  // ... rest of component unchanged (JSX same as current)
```

- [ ] **Step 3: Verify build**

```bash
cd /Users/cristianradu/Documents/programming/elaw-templating/frontend && npx tsc --noEmit 2>&1 | head -20
```

Expected: No TypeScript errors

- [ ] **Step 4: Commit**

```bash
git add frontend/package.json frontend/package-lock.json frontend/src/components/DropZone.tsx
git commit -m "feat: add JSZip-based 2-column DOCX detection in DropZone"
```

---

### Task 6: Frontend — UI State (OptionsForm + App)

**Files:**
- Modify: `frontend/src/components/OptionsForm.tsx` — add `disabled` prop
- Modify: `frontend/src/App.tsx` — add `isTwoColumn` state, wire everything
- Modify: `frontend/src/messages.ts` — add `convertToCellColumns` string

- [ ] **Step 1: Update messages.ts**

Add to `Messages` interface:
```typescript
app: {
  // ... existing
  convertToCellColumns: string
}
```

In `en`:
```typescript
app: {
  // ...
  convertToCellColumns: 'Convert to cell columns',
}
```

In `ro`:
```typescript
app: {
  // ...
  convertToCellColumns: 'Convertește în coloane',
}
```

- [ ] **Step 2: Update OptionsForm.tsx**

Add `disabled?: boolean` prop. When true, disable lang and mode selects:

```tsx
interface Props {
  // ... existing props
  disabled?: boolean
}

export default function OptionsForm({
  lang, mode, providerName, modelName,
  onLangChange, onModeChange, onProviderChange, onModelChange,
  disabled = false,
}: Props) {
  // ...

  return (
    <div className="space-y-4">
      <div className="grid grid-cols-2 gap-4">
        <div>
          <label className="block text-sm font-medium text-brand-500 mb-1">{msg.translateTo}</label>
          <select
            value={lang}
            onChange={(e) => onLangChange(e.target.value)}
            disabled={disabled}
            className={`w-full border border-brand-200 rounded-lg px-3 py-2 text-sm focus:ring-2 focus:ring-brand-300 focus:border-brand-300 text-brand-500 ${
              disabled ? 'opacity-50 cursor-not-allowed' : ''
            }`}
          >
            <option value="ro">{msg.romanian}</option>
            <option value="en">{msg.english}</option>
          </select>
        </div>
        <div>
          <label className="block text-sm font-medium text-brand-500 mb-1">{msg.mode}</label>
          <select
            value={mode}
            onChange={(e) => onModeChange(e.target.value)}
            disabled={disabled}
            className={`w-full border border-brand-200 rounded-lg px-3 py-2 text-sm focus:ring-2 focus:ring-brand-300 focus:border-brand-300 text-brand-500 ${
              disabled ? 'opacity-50 cursor-not-allowed' : ''
            }`}
          >
            <option value="inline">{msg.inline}</option>
            <option value="side-by-side">{msg.sideBySide}</option>
          </select>
        </div>
      </div>
      {/* ... rest unchanged */}
    </div>
  )
}
```

- [ ] **Step 3: Update App.tsx**

Add `isTwoColumn` state, wire DropZone detection, change button text, adjust submit:

```tsx
function App() {
  // ... existing state
  const [isTwoColumn, setIsTwoColumn] = useState(false)

  // Reset isTwoColumn when file changes
  useEffect(() => {
    setIsTwoColumn(false)
  }, [file])

  const handleSubmit = useCallback(async () => {
    if (!file) return
    setStatus('uploading')
    setError('')

    const formData = new FormData()
    formData.append('file', file)
    if (isTwoColumn) {
      formData.append('transform2cell', 'true')
    } else {
      formData.append('lang', lang)
      formData.append('mode', mode)
    }
    if (providerName) formData.append('provider', providerName)
    if (modelName) formData.append('model', modelName)
    // ... rest of submit unchanged
  }, [file, lang, mode, providerName, modelName, isTwoColumn, msg])

  // In JSX, pass isTwoColumn:
  <DropZone file={file} onFile={setFile} disabled={status === 'uploading'} onTwoColumnDetected={setIsTwoColumn} />
  <OptionsForm
    lang={lang}
    mode={mode}
    providerName={providerName}
    modelName={modelName}
    onLangChange={setLang}
    onModeChange={setMode}
    onProviderChange={setProviderName}
    onModelChange={setModelName}
    disabled={isTwoColumn}
  />
  <button ...>
    {status === 'uploading' ? msg.uploading : (isTwoColumn ? msg.convertToCellColumns : msg.convert)}
  </button>
```

- [ ] **Step 4: Verify build**

```bash
cd /Users/cristianradu/Documents/programming/elaw-templating/frontend && npx tsc --noEmit 2>&1 | head -20
```

Expected: No TypeScript errors

- [ ] **Step 5: Build frontend**

```bash
cd /Users/cristianradu/Documents/programming/elaw-templating/frontend && npm run build
```

Expected: Build succeeds, outputs to `frontend/dist/`

- [ ] **Step 6: Commit**

```bash
git add frontend/src/components/OptionsForm.tsx frontend/src/App.tsx frontend/src/messages.ts frontend/dist/
git commit -m "feat: add 2-column UI state with disabled lang/mode and cell columns button"
```

---

### Task 7: Run All Tests + Final Verification

- [ ] **Step 1: Run all existing tests**

```bash
cd /Users/cristianradu/Documents/programming/elaw-templating && python3 translate_docx.py
```

Expected: All tests PASS

- [ ] **Step 2: Run new transform2cell smoke test**

```bash
cd /Users/cristianradu/Documents/programming/elaw-templating && python3 -c "
import sys; sys.path.insert(0, '.')
if 'translate_docx' in sys.modules:
    del sys.modules['translate_docx']
import translate_docx
import tempfile, os
path = os.path.join(tempfile.mkdtemp(), 'test.docx')
translate_docx._create_2column_test_docx(path)
out = path.replace('.docx', '_2cell.docx')
translate_docx.transform2cell(path, out)
from docx import Document
d = Document(out)
assert len(d.tables) == 1
assert len(d.tables[0].rows) == 3
assert 'First original' in d.tables[0].rows[0].cells[0].text
assert 'Primul paragraf' in d.tables[0].rows[0].cells[1].text
os.unlink(path); os.unlink(out)
print('Smoke test PASS')
"
```

- [ ] **Step 3: Test CLI entry point**

```bash
cd /Users/cristianradu/Documents/programming/elaw-templating && python3 -c "
import sys; sys.path.insert(0, '.')
if 'translate_docx' in sys.modules:
    del sys.modules['translate_docx']
import translate_docx
args = translate_docx.parse_args(['input.docx', '--transform2cell'])
assert args.transform2cell
assert args.input == 'input.docx'
assert args.lang is None  # not required
print('CLI test PASS')
"
```

- [ ] **Step 4: Final commit if needed**

```bash
git add -A
git commit -m "chore: final cleanup and test fixes"
```
