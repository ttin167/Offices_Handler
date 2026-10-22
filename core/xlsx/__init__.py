"""
core.xlsx — Enterprise Excel OpenPyXL reader, writer, template mutator, streaming engine, and 12-Gate validator.
"""

from __future__ import annotations

from .reader import read_xlsx, read_xlsx_to_json
from .writer import (
    UniversalExcelEngine,
    write_xlsx,
    write_xlsx_from_json_file,
    mutate_template_excel,
    safe_set_cell,
    shift_formula_string,
    apply_worksheet_print_setup,
    sanitize_xml_string,
    parse_vietnamese_date,
    parse_accounting_and_localized_number,
    normalize_modern_formula,
)
from .streaming import (
    stream_read_xlsx,
    stream_write_xlsx,
    estimate_xlsx_metrics,
    STREAMING_FILE_SIZE_THRESHOLD_MB,
    STREAMING_ROW_COUNT_THRESHOLD,
)
from .validator import validate_excel_sheet
from .formula_playbook import (
    get_formula_catalog,
    generate_formula_playbook_xlsx,
    FORMULA_CATALOG,
)

__all__ = [
    "UniversalExcelEngine",
    "read_xlsx",
    "read_xlsx_to_json",
    "write_xlsx",
    "write_xlsx_from_json_file",
    "mutate_template_excel",
    "safe_set_cell",
    "shift_formula_string",
    "apply_worksheet_print_setup",
    "sanitize_xml_string",
    "parse_vietnamese_date",
    "parse_accounting_and_localized_number",
    "normalize_modern_formula",
    "stream_read_xlsx",
    "stream_write_xlsx",
    "estimate_xlsx_metrics",
    "STREAMING_FILE_SIZE_THRESHOLD_MB",
    "STREAMING_ROW_COUNT_THRESHOLD",
    "validate_excel_sheet",
    "get_formula_catalog",
    "generate_formula_playbook_xlsx",
    "FORMULA_CATALOG",
]


