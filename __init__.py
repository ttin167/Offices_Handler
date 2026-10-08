"""
Universal Document Studio — PDF/DOCX/XLSX/PPTX/Markdown converter + AI editing toolkit.
"""
try:
    from converter_engine import convert_file, parse_page_range, get_pdf_page_count
    from gui import launch_gui
    from cli import run_cli
    from docx_reader import read_docx, read_docx_to_json
    from docx_writer import write_docx, write_docx_from_json_file
    from xlsx_reader import read_xlsx, read_xlsx_to_json
    from xlsx_writer import write_xlsx, write_xlsx_from_json_file
    from docx_advanced_engine import (
        inject_dynamic_page_numbers,
        format_figure_captions,
        rebuild_table_of_contents,
        sanitize_emojis_and_symbols,
        translate_vn_to_en_protected,
        safe_save_docx,
    )
    from pptx_writer import (
        write_pptx_from_spec,
        add_title_slide,
        add_content_slide,
        add_diagram_slide,
        add_two_column_slide,
        add_kpi_cards_slide,
        add_table_slide,
        safe_save_pptx,
    )
    from pptx_reader import read_pptx, read_pptx_to_json
    from pptx_engine import (
        compile_markdown_file_to_pptx,
        parse_markdown_to_slides_spec,
        insert_diagram_into_pptx,
    )
except (ImportError, ValueError):
    from .converter_engine import convert_file, parse_page_range, get_pdf_page_count  # type: ignore
    from .gui import launch_gui  # type: ignore
    from .cli import run_cli  # type: ignore
    from .docx_reader import read_docx, read_docx_to_json  # type: ignore
    from .docx_writer import write_docx, write_docx_from_json_file  # type: ignore
    from .xlsx_reader import read_xlsx, read_xlsx_to_json  # type: ignore
    from .xlsx_writer import write_xlsx, write_xlsx_from_json_file  # type: ignore
    from .docx_advanced_engine import (  # type: ignore
        inject_dynamic_page_numbers,
        format_figure_captions,
        rebuild_table_of_contents,
        sanitize_emojis_and_symbols,
        translate_vn_to_en_protected,
        safe_save_docx,
    )
    from .pptx_writer import (  # type: ignore
        write_pptx_from_spec,
        add_title_slide,
        add_content_slide,
        add_diagram_slide,
        add_two_column_slide,
        add_kpi_cards_slide,
        add_table_slide,
        safe_save_pptx,
    )
    from .pptx_reader import read_pptx, read_pptx_to_json  # type: ignore
    from .pptx_engine import (  # type: ignore
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
