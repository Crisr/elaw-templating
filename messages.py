MESSAGES = {
    "cli_desc": "Translate DOCX documents using LLM",
    "cli_input_help": "Path to input DOCX file",
    "cli_lang_help": "Target language (ro=Romanian, en=English)",
    "cli_mode_help": "Output layout mode",
    "cli_output_help": "Output path (auto-generated if omitted)",
    "cli_provider_help": "Provider key from config.json",
    "cli_model_help": "Override model name",
    "cli_config_help": "Config file path",
    "provider_not_found": "Provider '{name}' not found in config. Available: {providers}",
    "file_not_found": "File not found: {path}",
    "err_translate_chunk": "Error translating chunk {i}: {e}",
    "warn_chunks_failed": "Warning: {count} chunk(s) failed: {indices}",
    "saved": "Translated document saved to: {output}",
    "test_pass": "  {name} PASS",
    "all_pass": "All tests PASS",
    "llm_system_prompt": (
        "You are a legal document translator. Translate the following paragraphs "
        "to {lang}. Preserve all paragraph IDs exactly ([P0], [P1], ...). "
        "Preserve {{...}} placeholders without translating them. "
        "Output ONLY the translated paragraphs with their IDs \u2014 no extra text."
    ),
}
