# Transform2Cell — Convert Word 2-Column Layout to Table

> **Spec for the `--transform2cell` feature**

## Goal

Add `--transform2cell` flag to `translate_docx.py` that converts a Word 2-column layout document into a table-based 2-column document (original left, translation right). Update the React frontend to detect 2-column DOCX files on drop and adjust UI (disable lang/mode, change button text).

## Architecture

Three-tier fallback: try exact column-break marker first, heuristic midpoint split + AI verification second, full AI matching third, AI re-translation last. The frontend uses JSZip to read DOCX structure client-side and detect columns before upload.

## Algorithm: Column Split Detection (Tiered)

1. **Exact:** Scan `<w:br w:type="column"/>` in paragraph XML → exact split index
2. **Heuristic:** `len(paras) // 2` when no break marker found
3. **AI verify:** Single LLM call verifies/adjusts the pairing
4. **AI full match:** LLM re-assigns every paragraph to a column + pairs them (fallback)
5. **Re-translate:** Reuse `translate_all` (last resort)

Once split index is known, pair by position: col1[i] ↔ col2[i] by order.

## Output Document

Invisible borderless table (same pattern as existing `side-by-side` mode). Page dimensions copied from source. Run formatting preserved. Numbering labels applied (to render correctly in table cells, matching existing `write_side_by_side` behavior).

## CLI

```
--transform2cel              Standalone flag, no --lang/--mode required
--provider, --model          Optional — for AI verification step
--config, --output           Same as existing
```

When `--transform2cell` is set without `--provider`, skip AI verification (heuristic only).

## Backend

- `server.py`: Accept `transform2cell: bool = Form(False)` on `/api/translate`. Skip `lang`/`mode` validation when `True`. Store mode as `"transform2cell"` in DB.
- `worker.py`: When `mode == "transform2cell"`, call `translate_docx.transform2cell(source_path, result_path, provider)`.
- Download filename: `transformed_cell.docx`.

## Frontend

- **JSZip** (new dependency): Read `word/document.xml` on file drop, detect `<w:cols w:num="2">`
- **DropZone**: New `onTwoColumnDetected` callback
- **OptionsForm**: `disabled` prop — when true, grey out "Translate to" and "Mode" selects
- **App.tsx**: `isTwoColumn` state — changes button text to "Convert to cell columns", skips `lang`/`mode` in FormData, adds `transform2cell: "true"`

## Files Modified

| File | Changes |
|---|---|
| `translate_docx.py` | +8 functions, modify `parse_args()`, `main()`, messages |
| `server.py` | Add `transform2cell` form field |
| `worker.py` | Handle `mode == "transform2cell"` |
| `frontend/package.json` | Add `jszip` |
| `frontend/src/App.tsx` | `isTwoColumn` state, button text, submit logic |
| `frontend/src/components/DropZone.tsx` | JSZip inspection, callback |
| `frontend/src/components/OptionsForm.tsx` | `disabled` prop |
| `frontend/src/messages.ts` | New string: `convertToCellColumns` |

## Testing

Programmatically generate synthetic 2-column `.docx` fixtures. Unit tests for detection, split, pairing, table writing. Integration test with AI verification skipped (mock).
