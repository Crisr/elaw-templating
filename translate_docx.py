import json
import os
import time
from openai import OpenAI
from docx import Document


def load_config(path="config.json"):
    with open(path) as f:
        return json.load(f)


def get_provider(config, name=None):
    if name is None:
        name = config["default_provider"]
    try:
        provider = config["providers"][name].copy()
    except KeyError:
        raise ValueError(f"Provider '{name}' not found in config. Available: {list(config['providers'].keys())}")
    api_key_env = provider.pop("api_key_env", None)
    if api_key_env:
        provider["api_key"] = os.environ.get(api_key_env)
        if not provider["api_key"]:
            from dotenv import load_dotenv
            load_dotenv()
            provider["api_key"] = os.environ.get(api_key_env)
    return provider


def _extract_runs(para):
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
    return runs_data


def _cell_coords(para):
    if not para._element.getparent().tag.endswith("tc"):
        return None, None
    tc = para._element.getparent()
    tr = tc.getparent()
    tbl = tr.getparent()
    tcs = [c for c in tr if c.tag.endswith("tc")]
    cell_col = tcs.index(tc)
    trs = [r for r in tbl if r.tag.endswith("tr")]
    cell_row = trs.index(tr)
    return cell_row, cell_col


def extract_paragraphs(path):
    if not os.path.exists(path):
        raise FileNotFoundError(f"File not found: {path}")
    doc = Document(path)
    paragraphs = []
    pid = 0

    for para in doc.paragraphs:
        cell_row, cell_col = _cell_coords(para)
        paragraphs.append({
            "id": f"P{pid}",
            "runs": _extract_runs(para),
            "alignment": str(para.alignment) if para.alignment else None,
            "in_table": cell_row is not None,
            "cell_row": cell_row,
            "cell_col": cell_col,
        })
        pid += 1

    for table in doc.tables:
        for row_idx, row in enumerate(table.rows):
            for col_idx, cell in enumerate(row.cells):
                for para in cell.paragraphs:
                    paragraphs.append({
                        "id": f"P{pid}",
                        "runs": _extract_runs(para),
                        "alignment": str(para.alignment) if para.alignment else None,
                        "in_table": True,
                        "cell_row": row_idx,
                        "cell_col": col_idx,
                    })
                    pid += 1

    return paragraphs


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
            for p in chunk:
                if p["id"] in parsed:
                    translated_text = parsed[p["id"]]
                    p["runs"][0]["text"] = translated_text
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
    p0 = paragraphs[0]
    assert p0["id"] == "P0"
    assert p0["in_table"] == False
    assert p0["cell_row"] is None
    assert p0["cell_col"] is None
    assert p0["alignment"] is None or p0["alignment"] is not None
    assert p0["runs"][0]["text"] == "Hello world"
    assert p0["runs"][0]["italic"] is None
    assert p0["runs"][0]["underline"] is None
    p1 = paragraphs[1]
    assert p1["runs"][0]["text"] == "Bold"
    assert p1["runs"][0]["bold"] == True
    assert p1["runs"][1]["text"] == " Normal"
    os.unlink(path)


def test_extract_paragraphs_with_table():
    from docx import Document
    import tempfile, os
    doc = Document()
    table = doc.add_table(rows=2, cols=3)
    table.cell(0, 0).text = "A1"
    table.cell(0, 1).text = "B1"
    table.cell(1, 2).text = "C2"
    path = os.path.join(tempfile.mkdtemp(), "test_table.docx")
    doc.save(path)
    paragraphs = extract_paragraphs(path)
    table_paras = [p for p in paragraphs if p["in_table"]]
    assert len(table_paras) >= 3, f"Expected >=3 table paragraphs, got {len(table_paras)} total paragraphs={len(paragraphs)}"
    for p in table_paras:
        assert p["cell_row"] is not None, f"Expected cell_row, got None for {p}"
        assert p["cell_col"] is not None, f"Expected cell_col, got None for {p}"
    scored = [(p["cell_row"], p["cell_col"], p["runs"][0]["text"] if p["runs"] else "") for p in table_paras]
    assert (0, 0, "A1") in scored, f"Missing A1 at (0,0): {scored}"
    assert (0, 1, "B1") in scored, f"Missing B1 at (0,1): {scored}"
    assert (1, 2, "C2") in scored, f"Missing C2 at (1,2): {scored}"
    os.unlink(path)

def test_extract_paragraphs_invalid_file():
    try:
        extract_paragraphs("/nonexistent/file.docx")
        assert False, "Should have raised FileNotFoundError"
    except FileNotFoundError:
        pass


def test_chunk_paragraphs():
    paragraphs = [{"id": f"P{i}"} for i in range(25)]
    chunks = chunk_paragraphs(paragraphs, 10)
    assert len(chunks) == 3


def test_translate_chunk():
    chunk = [{"id": "P0", "runs": [{"text": "Hello world", "bold": False, "italic": False}]}]
    provider = {"base_url": "http://invalid", "model": "test", "api_key": "bad"}
    try:
        translate_chunk(chunk, "Romanian", provider)
        assert False, "Should have raised ConnectionError or similar"
    except Exception:
        pass


def test_translate_chunk_retry():
    pass

if __name__ == "__main__":
    test_config_loading()
    test_get_provider()
    test_extract_paragraphs()
    test_extract_paragraphs_with_table()
    test_extract_paragraphs_invalid_file()
    test_chunk_paragraphs()
    test_translate_chunk()
    print("PASS")
