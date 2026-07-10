MESSAGES = {
    "cli_desc": "Translate DOCX documents using LLM",
    "cli_input_help": "Path to input DOCX file",
    "cli_lang_help": "Target language (ro=Romanian, en=English)",
    "cli_mode_help": "Output layout mode",
    "cli_output_help": "Output path (auto-generated if omitted)",
    "cli_provider_help": "Provider key from config.json",
    "cli_model_help": "Override model name",
    "cli_config_help": "Config file path",
    "cli_concurrency_help": "Parallel chunks (0=auto-detect, default: 0)",
    "detecting_concurrency": "Detecting optimal concurrency... {n}",
    "provider_not_found": "Provider '{name}' not found in config. Available: {providers}",
    "file_not_found": "File not found: {path}",
    "progress": "Translating: [{bar}] {pct:.0f}% ({done}/{total} chunks)",
    "err_translate_chunk": "Error translating chunk {i}: {e}",
    "warn_chunks_failed": "Warning: {count} chunk(s) failed: {indices}",
    "saved": "Translated document saved to: {output}",
    "test_pass": "  {name} PASS",
    "all_pass": "All tests PASS",
    "err_invalid_lang": "lang must be 'ro' or 'en'",
    "err_invalid_mode": "mode must be 'inline' or 'side-by-side'",
    "err_not_docx": "File is not a valid DOCX (missing ZIP header)",
    "err_job_not_found": "Job not found",
    "err_not_done": "Translation not yet complete",
    "err_result_cleaned": "Result file has been cleaned up",
    "llm_system_prompt": (
        "You are a legal document translator. Translate the following paragraphs "
        "to {lang}. Preserve all paragraph IDs exactly ([P0], [P1], ...). "
        "Preserve {{...}} placeholders without translating them. "
        "Output ONLY the translated paragraphs with their IDs \u2014 no extra text."
    ),
    "warn_no_2column_layout": "Warning: Input document does not appear to have a 2-column layout. Results may be unexpected.",
    "info_heuristic_column_split": "Info: No column break marker found. Using heuristic split at paragraph {idx}.",
    "warn_transform2cell_mismatch": "Warning: Column paragraph counts differ (col1={c1}, col2={c2}).",
    "cli_transform2cell_help": "Convert 2-column Word layout to table-based document",
}
