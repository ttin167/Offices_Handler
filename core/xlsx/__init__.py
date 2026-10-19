"""
core.xlsx — Excel OpenPyXL reader, writer, template mutator, and 12-Gate validator.
"""

from __future__ import annotations

from .reader import read_xlsx, read_xlsx_to_json
from .writer import (
    write_xlsx,
    write_xlsx_from_json_file,
    mutate_template_excel,
    apply_worksheet_print_setup,
)
from .validator import validate_excel_sheet

__all__ = [
    "read_xlsx",
    "read_xlsx_to_json",
    "write_xlsx",
    "write_xlsx_from_json_file",
    "mutate_template_excel",
    "apply_worksheet_print_setup",
    "validate_excel_sheet",
]
