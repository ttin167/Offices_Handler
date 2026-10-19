"""
core.pptx — Widescreen 16:9 Presentation Studio, slide engines, themes, mockups, and converters.
"""

from __future__ import annotations

from .reader import read_pptx, read_pptx_to_json
from .writer import (
    write_pptx_from_spec,
    add_title_slide,
    add_content_slide,
    add_diagram_slide,
    add_two_column_slide,
    add_kpi_cards_slide,
    add_table_slide,
    add_split_diagram_slide,
    add_mockup_slide,
    add_split_mockup_slide,
    add_kpi_summary_slide,
    add_timeline_roadmap_slide,
    add_swot_matrix_slide,
    add_process_chevron_slide,
    safe_save_pptx,
    get_theme,
    THEMES,
)
from .charts import (
    add_chart_to_slide,
    add_chart_slide,
    build_chart_data,
)
from .engine import (
    compile_markdown_file_to_pptx,
    compile_markdown_to_presentation,
    parse_markdown_to_slides_spec,
    insert_diagram_into_pptx,
)
from .math import add_math_text_to_paragraph, parse_inline_math
from .docx_to_pptx import convert_docx_to_pptx, convert_pdf_to_pptx

__all__ = [
    "read_pptx",
    "read_pptx_to_json",
    "write_pptx_from_spec",
    "add_title_slide",
    "add_content_slide",
    "add_diagram_slide",
    "add_two_column_slide",
    "add_kpi_cards_slide",
    "add_table_slide",
    "add_split_diagram_slide",
    "add_mockup_slide",
    "add_split_mockup_slide",
    "add_kpi_summary_slide",
    "add_timeline_roadmap_slide",
    "add_swot_matrix_slide",
    "add_process_chevron_slide",
    "add_chart_to_slide",
    "add_chart_slide",
    "build_chart_data",
    "safe_save_pptx",
    "get_theme",
    "THEMES",
    "compile_markdown_file_to_pptx",
    "compile_markdown_to_presentation",
    "parse_markdown_to_slides_spec",
    "insert_diagram_into_pptx",
    "add_math_text_to_paragraph",
    "parse_inline_math",
    "convert_docx_to_pptx",
    "convert_pdf_to_pptx",
]
