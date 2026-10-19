"""
core.docx — Advanced Word OpenXML processing, reading, writing, styling, and patching.
"""

from __future__ import annotations

from .reader import read_docx, read_docx_to_json
from .writer import (
    write_docx,
    write_docx_from_json_file,
    patch_docx_template,
    inject_content_into_docx,
    inject_diagram_into_docx,
)
from .presets import (
    apply_preset_to_document,
    load_preset,
    list_available_presets,
    build_administrative_header,
)
from .advanced_engine import (
    inject_dynamic_page_numbers,
    format_figure_captions,
    rebuild_table_of_contents,
    sanitize_emojis_and_symbols,
    translate_vn_to_en_protected,
    safe_save_docx,
    VietnameseTypographySanitizer,
    sanitize_vietnamese_typography,
)
from .lists import ListContextTracker, apply_list_formatting, ensure_numbering_definitions
from .math import add_math_to_paragraph
from .mail_merge import (
    inspect_template_tokens,
    evaluate_condition,
    resolve_conditional_text,
    read_data_source,
    execute_mail_merge,
)
from .diff_engine import (
    RedlineDiffEngine,
    TrackChangeIdGenerator,
    tokenize_text,
    diff_token_sequences,
    compare_docx_documents,
)
from .post_processor import post_process_docx, calculate_heuristic_column_widths

__all__ = [
    "read_docx",
    "read_docx_to_json",
    "write_docx",
    "write_docx_from_json_file",
    "patch_docx_template",
    "inject_content_into_docx",
    "inject_diagram_into_docx",
    "inject_dynamic_page_numbers",
    "format_figure_captions",
    "rebuild_table_of_contents",
    "sanitize_emojis_and_symbols",
    "translate_vn_to_en_protected",
    "safe_save_docx",
    "VietnameseTypographySanitizer",
    "sanitize_vietnamese_typography",
    "apply_preset_to_document",
    "load_preset",
    "list_available_presets",
    "build_administrative_header",
    "ListContextTracker",
    "apply_list_formatting",
    "ensure_numbering_definitions",
    "add_math_to_paragraph",
    "post_process_docx",
    "calculate_heuristic_column_widths",
    "inspect_template_tokens",
    "evaluate_condition",
    "resolve_conditional_text",
    "read_data_source",
    "execute_mail_merge",
    "RedlineDiffEngine",
    "TrackChangeIdGenerator",
    "tokenize_text",
    "diff_token_sequences",
    "compare_docx_documents",
]
