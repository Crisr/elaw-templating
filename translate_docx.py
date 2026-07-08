import json
import os
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

if __name__ == "__main__":
    test_config_loading()
    test_get_provider()
    test_extract_paragraphs()
    test_extract_paragraphs_with_table()
    test_extract_paragraphs_invalid_file()
    print("PASS")
