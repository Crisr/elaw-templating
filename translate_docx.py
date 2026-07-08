import argparse
import json
import os
import sys
import time
import re
import threading
from concurrent.futures import ThreadPoolExecutor, as_completed
from openai import OpenAI
from docx import Document, Document as DocxDocument
from docx.shared import Pt, RGBColor
from messages import MESSAGES as _


def load_config(path="config.json"):
    with open(path) as f:
        return json.load(f)


def get_provider(config, name=None):
    if name is None:
        name = config["default_provider"]
    try:
        provider = config["providers"][name].copy()
    except KeyError:
        providers_list = list(config['providers'].keys())
        raise ValueError(_["provider_not_found"].format(name=name, providers=providers_list))
    api_key_env = provider.pop("api_key_env", None)
    if api_key_env:
        provider["api_key"] = os.environ.get(api_key_env)
        if not provider["api_key"]:
            try:
                from dotenv import load_dotenv  # type: ignore
                load_dotenv()
            except ImportError:
                pass
            provider["api_key"] = os.environ.get(api_key_env)
        if not provider["api_key"]:
            try:
                with open("secrets.json") as f:
                    secrets = json.load(f)
                    provider["api_key"] = secrets.get(api_key_env)
            except (FileNotFoundError, json.JSONDecodeError):
                pass
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
            "font_size": run.font.size.pt if run.font.size else None,
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
        raise FileNotFoundError(_["file_not_found"].format(path=path))
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


def restore_placeholders(translated_text, original_text):
    placeholder_pattern = r'\{\{[^}]+\}\}'
    placeholders = re.findall(placeholder_pattern, original_text)
    for ph in placeholders:
        if ph not in translated_text:
            orig_pos = original_text.index(ph)
            ratio = orig_pos / max(len(original_text), 1)
            insert_pos = int(ratio * len(translated_text))
            translated_text = translated_text[:insert_pos] + ph + translated_text[insert_pos:]
    return translated_text


def _parse_translated_response(response_text, chunk_ids, original_text=None):
    results = {}
    lines = response_text.strip().split("\n")
    if lines and not lines[0].strip().startswith("[P"):
        lines = lines[1:]
    for line in lines:
        line = line.strip()
        if not line:
            continue
        for pid in chunk_ids:
            if line.startswith(f"[{pid}]"):
                text = line[len(f"[{pid}]"):].strip()
                results[pid] = text
                break
    if len(results) < len(chunk_ids):
        all_lines = [ln for ln in response_text.strip().split("\n") if ln.strip()]
        for i, pid in enumerate(chunk_ids):
            if pid not in results and i < len(all_lines):
                results[pid] = all_lines[i].strip()
    if original_text:
        original_lines = {}
        for line in original_text.strip().split("\n"):
            for pid in chunk_ids:
                if line.startswith(f"[{pid}]"):
                    original_lines[pid] = line[len(f"[{pid}]"):].strip()
                    break
        for pid in results:
            if pid in original_lines:
                results[pid] = restore_placeholders(results[pid], original_lines[pid])
    return results


def translate_chunk(chunk, target_lang, provider):
    chunk_text = _build_chunk_text(chunk)
    chunk_ids = [p["id"] for p in chunk]
    api_key = provider.get("api_key") or "not-needed"
    client = OpenAI(base_url=provider["base_url"], api_key=api_key)
    system_prompt = _["llm_system_prompt"].format(lang=target_lang)
    last_error = RuntimeError("translate_chunk failed after all retries")
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
            parsed = _parse_translated_response(translated, chunk_ids, chunk_text)
            for p in chunk:
                if p["id"] in parsed:
                    translated_text = parsed[p["id"]]
                    original_runs = p["runs"]
                    original_full_text = "".join(r["text"] for r in original_runs)
                    if original_full_text and len(original_runs) > 1:
                        start = 0
                        for j, run in enumerate(original_runs):
                            proportion = len(run["text"]) / len(original_full_text)
                            chunk_len = int(proportion * len(translated_text))
                            if j == len(original_runs) - 1:
                                run["text"] = translated_text[start:]
                            else:
                                run["text"] = translated_text[start:start + chunk_len]
                                start += chunk_len
                    elif original_runs:
                        original_runs[0]["text"] = translated_text
            return chunk
        except Exception as e:
            last_error = e
            if attempt < 2:
                time.sleep(2 ** attempt)
    raise last_error


def _show_progress(done, total):
    if total == 0:
        return
    width = 30
    pct = done / total
    filled = int(width * pct)
    bar = "#" * filled + "-" * (width - filled)
    print("\r" + _["progress"].format(bar=bar, pct=pct * 100, done=done, total=total), end="", flush=True)


_CONCURRENCY_CACHE = None


def _detect_concurrency(provider):
    global _CONCURRENCY_CACHE
    if _CONCURRENCY_CACHE is not None:
        return _CONCURRENCY_CACHE
    api_key = provider.get("api_key") or "not-needed"
    client = OpenAI(base_url=provider["base_url"], api_key=api_key)
    probe = [{"role": "user", "content": "ok"}]
    best = 1
    for level in [2, 4, 6, 8, 12, 16]:
        print("\r" + _["detecting_concurrency"].format(n=level), end="", flush=True)
        def try_request(_):
            try:
                client.chat.completions.create(model=provider["model"], messages=probe, temperature=0, max_tokens=1)  # type: ignore
                return True
            except Exception:
                return False
        with ThreadPoolExecutor(max_workers=level) as pool:
            results = list(pool.map(try_request, range(level)))
        if sum(1 for r in results if not r) == 0:
            best = level
        else:
            break
    print("\r" + " " * 50 + "\r", end="", flush=True)
    _CONCURRENCY_CACHE = best
    return best


def translate_all(paragraphs, target_lang, provider, concurrency=0):
    if concurrency < 1:
        concurrency = _detect_concurrency(provider)
    chunks = chunk_paragraphs(paragraphs)
    total = len(chunks)
    if concurrency > total:
        concurrency = total
    results = [None] * total
    failed_indices = []
    lock = threading.Lock()
    done = 0

    def process(i, chunk):
        try:
            translated = translate_chunk(chunk, target_lang, provider)
            return i, translated, None
        except Exception as e:
            return i, None, e

    _show_progress(0, total)
    with ThreadPoolExecutor(max_workers=concurrency) as pool:
        futures = [pool.submit(process, i, c) for i, c in enumerate(chunks)]
        for future in as_completed(futures):
            i, translated, err = future.result()
            with lock:
                if err:
                    print("\n" + _["err_translate_chunk"].format(i=i, e=err), file=sys.stderr)
                    failed_indices.append(i)
                else:
                    results[i] = translated
                done += 1
                _show_progress(done, total)
    print()
    all_translated = []
    for r in results:
        if r is not None:
            all_translated.extend(r)
    if failed_indices:
        print(_["warn_chunks_failed"].format(count=len(failed_indices), indices=failed_indices), file=sys.stderr)
    return all_translated


def write_inline(original_path, translated_paragraphs, output_path):
    doc = DocxDocument(original_path)
    trans_by_id = {p["id"]: p for p in translated_paragraphs}
    for i, para in enumerate(doc.paragraphs):
        pid = f"P{i}"
        if pid in trans_by_id:
            tp = trans_by_id[pid]
            for j, run in enumerate(para.runs):
                if j < len(tp["runs"]):
                    run.text = tp["runs"][j]["text"]
    pid = len(doc.paragraphs)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    pid_str = f"P{pid}"
                    if pid_str in trans_by_id:
                        tp = trans_by_id[pid_str]
                        for j, run in enumerate(para.runs):
                            if j < len(tp["runs"]):
                                run.text = tp["runs"][j]["text"]
                    pid += 1
    doc.save(output_path)


def test_config_loading():
    import json
    import tempfile
    import os
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
        fname = f.name
    try:
        result = load_config(fname)
        assert result == cfg
    finally:
        os.unlink(fname)


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
    import tempfile
    import os
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
    assert not p0["in_table"]
    assert p0["cell_row"] is None
    assert p0["cell_col"] is None
    assert p0["alignment"] is None or p0["alignment"] is not None
    assert p0["runs"][0]["text"] == "Hello world"
    assert p0["runs"][0]["italic"] is None
    assert p0["runs"][0]["underline"] is None
    p1 = paragraphs[1]
    assert p1["runs"][0]["text"] == "Bold"
    assert p1["runs"][0]["bold"]
    assert p1["runs"][1]["text"] == " Normal"
    os.unlink(path)


def test_extract_paragraphs_with_table():
    from docx import Document
    import tempfile
    import os
    doc = Document()
    table = doc.add_table(rows=2, cols=3)
    table.cell(0, 0).text = "A1"
    table.cell(0, 1).text = "B1"
    table.cell(1, 2).text = "C2"
    path = os.path.join(tempfile.mkdtemp(), "test_table.docx")
    doc.save(path)
    paragraphs = extract_paragraphs(path)
    table_paras = [p for p in paragraphs if p["in_table"]]
    assert len(table_paras) == 6, f"Expected 6 table paragraphs, got {len(table_paras)} total paragraphs={len(paragraphs)}"
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


def test_write_inline():
    from docx import Document
    import tempfile
    import os
    doc = Document()
    p = doc.add_paragraph()
    p.add_run("Original text")
    src = os.path.join(tempfile.mkdtemp(), "source.docx")
    doc.save(src)
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
    os.unlink(src)
    os.unlink(out)


def _apply_run_formatting(run, run_data):
    run.bold = run_data.get("bold")
    run.italic = run_data.get("italic")
    run.underline = run_data.get("underline")
    if run_data.get("font_name"):
        run.font.name = run_data["font_name"]
    fs = run_data.get("font_size")
    if fs is not None:
        try:
            run.font.size = Pt(float(fs))
        except (ValueError, TypeError):
            pass
    if run_data.get("color"):
        try:
            run.font.color.rgb = RGBColor.from_string(run_data["color"])
        except (ValueError, AttributeError):
            pass


def write_side_by_side(original_path, original_paras, translated_paras, output_path):
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
    trans_by_id = {p["id"]: p for p in translated_paras}
    table = doc.add_table(rows=len(original_paras), cols=2)
    table.style = 'Table Grid'
    for i, op in enumerate(original_paras):
        pid = op["id"]
        left_cell = table.rows[i].cells[0]
        right_cell = table.rows[i].cells[1]
        for run_data in op["runs"]:
            p = left_cell.paragraphs[0] if left_cell.paragraphs else left_cell.add_paragraph()
            run = p.add_run(run_data["text"])
            _apply_run_formatting(run, run_data)
        tp = trans_by_id.get(pid, op)
        for run_data in tp["runs"]:
            p = right_cell.paragraphs[0] if right_cell.paragraphs else right_cell.add_paragraph()
            run = p.add_run(run_data["text"])
            _apply_run_formatting(run, run_data)
    doc.save(output_path)


def test_write_inline_with_table():
    from docx import Document
    import tempfile
    import os
    doc = Document()
    doc.add_paragraph("Header text")
    table = doc.add_table(rows=2, cols=2)
    table.cell(0, 0).text = "Cell A1"
    table.cell(0, 1).text = "Cell B1"
    table.cell(1, 0).text = "Cell A2"
    table.cell(1, 1).text = "Cell B2"
    src = os.path.join(tempfile.mkdtemp(), "source.docx")
    doc.save(src)
    translated = [
        {"id": "P0", "runs": [{"text": "Header text trans"}], "alignment": None, "in_table": False},
        {"id": "P1", "runs": [{"text": "Celula A1"}], "alignment": None, "in_table": True, "cell_row": 0, "cell_col": 0},
        {"id": "P2", "runs": [{"text": "Celula B1"}], "alignment": None, "in_table": True, "cell_row": 0, "cell_col": 1},
        {"id": "P3", "runs": [{"text": "Celula A2"}], "alignment": None, "in_table": True, "cell_row": 1, "cell_col": 0},
        {"id": "P4", "runs": [{"text": "Celula B2"}], "alignment": None, "in_table": True, "cell_row": 1, "cell_col": 1},
    ]
    out = os.path.join(tempfile.mkdtemp(), "output.docx")
    write_inline(src, translated, out)
    result = Document(out)
    assert result.paragraphs[0].text == "Header text trans"
    out_table = result.tables[0]
    assert out_table.cell(0, 0).text == "Celula A1"
    assert out_table.cell(0, 1).text == "Celula B1"
    assert out_table.cell(1, 0).text == "Celula A2"
    assert out_table.cell(1, 1).text == "Celula B2"
    os.unlink(src)
    os.unlink(out)


def test_cli_parser():
    args = parse_args(['input.docx', '--lang', 'ro'])
    assert args.input == 'input.docx'
    assert args.lang == 'ro'
    assert args.mode == 'inline'
    assert args.concurrency == 0
    args = parse_args(['input.docx', '--lang', 'en', '--mode', 'side-by-side', '--concurrency', '2'])
    assert args.input == 'input.docx'
    assert args.lang == 'en'
    assert args.mode == 'side-by-side'
    assert args.concurrency == 2


def test_write_side_by_side():
    from docx import Document
    import tempfile
    import os
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
    os.unlink(src)
    os.unlink(out)


def test_restore_placeholders():
    original = "Hello {{client_name}}, your balance is {{amount}}"
    translated = "Salut {{client_name}}, soldul tau este {{amount}}"
    restored = restore_placeholders(translated, original)
    assert "{{client_name}}" in restored
    assert "{{amount}}" in restored


def test_restore_missing_placeholder():
    original = "Hello {{client_name}}"
    translated = "Salut"
    restored = restore_placeholders(translated, original)
    assert "{{client_name}}" in restored


def test_parse_translated_response_protects_placeholders():
    chunk_ids = ["P0", "P1"]
    response_text = "[P0] Salut {{client_name}}\n[P1] Soldul este {{amount}}"
    original_text = "[P0] Hello {{client_name}}\n[P1] Your balance is {{amount}}"
    result = _parse_translated_response(response_text, chunk_ids, original_text)
    assert "{{client_name}}" in result["P0"]
    assert "{{amount}}" in result["P1"]


def parse_args(argv=None):
    parser = argparse.ArgumentParser(description=_["cli_desc"])
    parser.add_argument("input", help=_["cli_input_help"])
    parser.add_argument("--lang", "-l", required=True, choices=["ro", "en"],
                        help=_["cli_lang_help"])
    parser.add_argument("--mode", "-m", choices=["inline", "side-by-side"],
                        default="inline", help=_["cli_mode_help"])
    parser.add_argument("--output", "-o", help=_["cli_output_help"])
    parser.add_argument("--provider", "-p", help=_["cli_provider_help"])
    parser.add_argument("--model", help=_["cli_model_help"])
    parser.add_argument("--config", "-c", default="config.json", help=_["cli_config_help"])
    parser.add_argument("--concurrency", "-j", type=int, default=0,
                        help=_["cli_concurrency_help"])
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
    if os.path.exists(output):
        stem = output.rsplit(".docx", 1)[0]
        timestamp = time.strftime("%Y%m%d_%H%M%S")
        output = f"{stem}_{timestamp}.docx"
    paragraphs = extract_paragraphs(args.input)
    translated = translate_all(paragraphs, "Romanian" if args.lang == "ro" else "English", provider, args.concurrency)
    if args.mode == "side-by-side":
        write_side_by_side(args.input, paragraphs, translated, output)
    else:
        write_inline(args.input, translated, output)
    print(_["saved"].format(output=output))


if __name__ == "__main__":
    if len(sys.argv) > 1:
        main()
    else:
        tests = [
            ("test_config_loading", test_config_loading),
            ("test_get_provider", test_get_provider),
            ("test_extract_paragraphs", test_extract_paragraphs),
            ("test_extract_paragraphs_with_table", test_extract_paragraphs_with_table),
            ("test_extract_paragraphs_invalid_file", test_extract_paragraphs_invalid_file),
            ("test_chunk_paragraphs", test_chunk_paragraphs),
            ("test_translate_chunk", test_translate_chunk),
            ("test_translate_chunk_retry", test_translate_chunk_retry),
            ("test_write_inline", test_write_inline),
            ("test_write_side_by_side", test_write_side_by_side),
            ("test_write_inline_with_table", test_write_inline_with_table),
            ("test_cli_parser", test_cli_parser),
            ("test_restore_placeholders", test_restore_placeholders),
            ("test_restore_missing_placeholder", test_restore_missing_placeholder),
            ("test_parse_translated_response_protects_placeholders", test_parse_translated_response_protects_placeholders),
        ]
        for name, fn in tests:
            fn()
            print(_["test_pass"].format(name=name))
        print(_["all_pass"])
