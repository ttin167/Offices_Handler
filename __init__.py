# -*- coding: utf-8 -*-
"""
Universal Document Studio — PDF/DOCX/XLSX/PPTX/Markdown converter + AI editing toolkit.
"""
from core.converter.engine import convert_file, parse_page_range, get_pdf_page_count
from ui.desktop_app import launch_gui
from cli.converter_cli import run_cli
from core.docx.reader import read_docx, read_docx_to_json
from core.docx.writer import write_docx, write_docx_from_json_file
from core.xlsx.reader import read_xlsx, read_xlsx_to_json
from core.xlsx.writer import write_xlsx, write_xlsx_from_json_file
from core.docx.advanced_engine import (
    inject_dynamic_page_numbers,
    format_figure_captions,
    rebuild_table_of_contents,
    sanitize_emojis_and_symbols,
    translate_vn_to_en_protected,
    safe_save_docx,
)
from core.pptx.writer import (
    write_pptx_from_spec,
    add_title_slide,
    add_content_slide,
    add_diagram_slide,
    add_two_column_slide,
    add_kpi_cards_slide,
    add_table_slide,
    safe_save_pptx,
)
from core.pptx.reader import read_pptx, read_pptx_to_json
from core.pptx.engine import (
    compile_markdown_file_to_pptx,
    parse_markdown_to_slides_spec,
    insert_diagram_into_pptx,
)

__all__ = [
    # Convert pipeline
    "convert_file", "parse_page_range", "get_pdf_page_count",
    "launch_gui", "run_cli",
    # AI editing toolkit — DOCX
    "read_docx", "read_docx_to_json",
    "write_docx", "write_docx_from_json_file",
    # Advanced DOCX Engine
    "inject_dynamic_page_numbers",
    "format_figure_captions",
    "rebuild_table_of_contents",
    "sanitize_emojis_and_symbols",
    "translate_vn_to_en_protected",
    "safe_save_docx",
    # AI editing toolkit — XLSX
    "read_xlsx", "read_xlsx_to_json",
    "write_xlsx", "write_xlsx_from_json_file",
    # AI editing toolkit — PPTX
    "write_pptx_from_spec",
    "add_title_slide",
    "add_content_slide",
    "add_diagram_slide",
    "add_two_column_slide",
    "add_kpi_cards_slide",
    "add_table_slide",
    "safe_save_pptx",
    "read_pptx",
    "read_pptx_to_json",
    "compile_markdown_file_to_pptx",
    "parse_markdown_to_slides_spec",
    "insert_diagram_into_pptx",
]
