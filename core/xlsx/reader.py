"""
xlsx_reader.py — XLSX → JSON Snapshot for AI editing.

Uses xlwings to read Excel files, which requires Microsoft Excel to be
installed. This allows reading computed formula values (not just formula
strings), which is essential for AI to understand actual data.

Captures per-cell:
  - address, value (computed), formula string, data_type
  - format: font, fill_color, bold, number_format
  - merged cell ranges per sheet
"""

from __future__ import annotations

import json
import os
from typing import Any


def _get_color_hex(xw_color) -> str | None:
    """Convert xlwings color tuple (R, G, B) or None to hex string."""
    if xw_color is None:
        return None
    try:
        r, g, b = int(xw_color[0]), int(xw_color[1]), int(xw_color[2])
        return f"#{r:02X}{g:02X}{b:02X}"
    except Exception:
        return None


def _col_index_to_letter(col_index: int) -> str:
    """Convert 1-based column index to Excel letter (1=A, 27=AA, ...)."""
    result = ""
    while col_index > 0:
        col_index, remainder = divmod(col_index - 1, 26)
        result = chr(65 + remainder) + result
    return result


def _serialise_sheet(sheet, include_range: str | None = None) -> dict[str, Any]:
    """
    Serialise a single xlwings Sheet to a dict.

    Args:
        sheet:         xlwings Sheet object.
        include_range: Optional Excel range string to limit output (e.g. 'A1:H20').
    """
    used = sheet.used_range
    if used is None:
        return {
            "sheet_name": sheet.name,
            "used_range": None,
            "cells": [],
            "merged_cells": [],
        }

    if include_range:
        rng = sheet.range(include_range)
    else:
        rng = used

    used_address = used.address.replace("$", "")

    cells_data = []
    for cell in rng:
        addr = cell.address.replace("$", "")
        value = cell.value
        formula_str = None

        # Read formula separately
        try:
            f = cell.formula
            if f and str(f).startswith("="):
                formula_str = f
        except Exception:
            pass

        # Data type
        if value is None:
            dtype = "empty"
        elif isinstance(value, bool):
            dtype = "boolean"
        elif isinstance(value, (int, float)):
            dtype = "number"
        elif isinstance(value, str):
            dtype = "string"
        else:
            # datetime or other
            dtype = type(value).__name__

        # Format attributes
        fmt: dict[str, Any] = {}
        try:
            font = cell.font
            fmt["font_name"]     = font.name
            fmt["font_size"]     = font.size
            fmt["bold"]          = font.bold
            fmt["italic"]        = font.italic
            fmt["underline"]     = font.underline
            fmt["font_color"]    = _get_color_hex(font.color)
        except Exception:
            pass
        try:
            fmt["fill_color"]    = _get_color_hex(cell.color)
        except Exception:
            pass
        try:
            fmt["number_format"] = cell.number_format
        except Exception:
            pass

        # Merged state
        merged = False
        try:
            merged = cell.merge_area.address != cell.address
        except Exception:
            pass

        cells_data.append({
            "address":   addr,
            "value":     value if not hasattr(value, "isoformat") else value.isoformat(),
            "formula":   formula_str,
            "data_type": dtype,
            "format":    fmt,
            "merged":    merged,
        })

    # Merged cell ranges
    merged_cells: list[str] = []
    try:
        for area in sheet.used_range.merge_areas:
            merged_cells.append(area.address.replace("$", ""))
    except Exception:
        pass

    return {
        "sheet_name":   sheet.name,
        "used_range":   used_address,
        "cells":        cells_data,
        "merged_cells": merged_cells,
    }


def _read_xlsx_openpyxl(
    xlsx_path: str,
    sheet_name: str | None = None,
    cell_range: str | None = None,
) -> dict[str, Any]:
    """
    Zero-dependency OpenPyXL fallback reader.
    Does not require Microsoft Excel to be installed, making it suitable
    for headless Linux servers, CI/CD pipelines, and systems without Office.
    """
    import openpyxl
    from openpyxl.cell.cell import MergedCell
    from openpyxl.utils import get_column_letter

    wb_form = openpyxl.load_workbook(xlsx_path, data_only=False)
    wb_val = None
    try:
        try:
            wb_val = openpyxl.load_workbook(xlsx_path, data_only=True)
        except Exception:
            wb_val = None

        target_sheets = [sheet_name] if sheet_name else wb_form.sheetnames
        sheets_data = []

        for sname in target_sheets:
            if sname not in wb_form.sheetnames:
                raise ValueError(f"Sheet '{sname}' not found. Available: {wb_form.sheetnames}")
            ws_form = wb_form[sname]
            ws_val = wb_val[sname] if wb_val and sname in wb_val.sheetnames else None

            merged_cells = [str(r.coord).replace("$", "") for r in ws_form.merged_cells.ranges]
            merged_coords_set = set()
            for m_range in ws_form.merged_cells.ranges:
                min_col, min_row, max_col, max_row = m_range.bounds
                for r in range(min_row, max_row + 1):
                    for c in range(min_col, max_col + 1):
                        if (r, c) != (min_row, min_col):
                            merged_coords_set.add((r, c))

            if cell_range:
                cells_matrix = ws_form[cell_range]
                if isinstance(cells_matrix, tuple):
                    if len(cells_matrix) > 0 and not isinstance(cells_matrix[0], (tuple, list)):
                        cell_list = list(cells_matrix)
                    else:
                        cell_list = [c for row in cells_matrix for c in row]
                else:
                    cell_list = [cells_matrix]
                used_address = cell_range.replace("$", "")
            else:
                cell_list = [c for row in ws_form.iter_rows() for c in row]
                max_c_letter = get_column_letter(ws_form.max_column) if ws_form.max_column else "A"
                used_address = f"A1:{max_c_letter}{ws_form.max_row or 1}"

            cells_data = []
            for c in cell_list:
                addr = c.coordinate
                val = None
                if ws_val:
                    try:
                        val = ws_val.cell(row=c.row, column=c.column).value
                    except Exception:
                        pass
                if val is None:
                    val = c.value

                formula_str = None
                raw_v = c.value
                if isinstance(raw_v, str) and raw_v.startswith("="):
                    formula_str = raw_v

                # Data type
                if val is None:
                    dtype = "empty"
                elif isinstance(val, bool):
                    dtype = "boolean"
                elif isinstance(val, (int, float)):
                    dtype = "number"
                elif isinstance(val, str):
                    dtype = "string"
                else:
                    dtype = type(val).__name__

                # Format
                fmt: dict[str, Any] = {}
                if c.font:
                    fmt["font_name"] = c.font.name
                    fmt["font_size"] = c.font.size
                    fmt["bold"] = bool(c.font.bold)
                    fmt["italic"] = bool(c.font.italic)
                    fmt["underline"] = bool(c.font.underline)
                    if c.font.color and c.font.color.rgb and isinstance(c.font.color.rgb, str):
                        fmt["font_color"] = f"#{c.font.color.rgb[-6:]}"
                if c.fill and getattr(c.fill, "fill_type", None) == "solid":
                    fc = getattr(c.fill, "fgColor", None) or getattr(c.fill, "start_color", None)
                    if fc and getattr(fc, "rgb", None) and isinstance(fc.rgb, str):
                        fmt["fill_color"] = f"#{fc.rgb[-6:]}"
                if c.number_format:
                    fmt["number_format"] = c.number_format

                is_merged = (c.row, c.column) in merged_coords_set or isinstance(c, MergedCell)

                cells_data.append({
                    "address": addr,
                    "value": val if not hasattr(val, "isoformat") else val.isoformat(),
                    "formula": formula_str,
                    "data_type": dtype,
                    "format": fmt,
                    "merged": is_merged,
                })

            sheets_data.append({
                "sheet_name": sname,
                "used_range": used_address,
                "cells": cells_data,
                "merged_cells": merged_cells,
            })

        return {
            "source_file": os.path.abspath(xlsx_path),
            "snapshot_version": "1.0",
            "sheets": sheets_data,
        }
    finally:
        try:
            wb_form.close()
        except Exception:
            pass
        if wb_val is not None:
            try:
                wb_val.close()
            except Exception:
                pass


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def read_xlsx(
    xlsx_path: str,
    sheet_name: str | None = None,
    cell_range: str | None = None,
    engine: str = "auto",
) -> dict[str, Any]:
    """
    Read an XLSX file and return a structured JSON-serialisable snapshot.

    Args:
        xlsx_path:  Absolute path to the .xlsx file.
        sheet_name: Optional sheet name to export (exports all sheets if None).
        cell_range: Optional cell range to limit export (e.g. 'A1:H20').
        engine:     'auto' (default: attempts xlwings, falls back to openpyxl),
                    'openpyxl' (pure headless, zero MS Excel dependency),
                    'xlwings' (requires Microsoft Excel Desktop installed).

    Returns:
        dict with keys: source_file, snapshot_version, sheets
    """
    if not os.path.isfile(xlsx_path):
        raise FileNotFoundError(f"File not found: {xlsx_path}")

    selected_engine = (engine or "auto").lower()

    if selected_engine == "openpyxl":
        return _read_xlsx_openpyxl(xlsx_path, sheet_name=sheet_name, cell_range=cell_range)

    # Attempt xlwings (either auto or explicit xlwings)
    try:
        import xlwings as xw

        app = xw.App(visible=False, add_book=False)
        try:
            wb = app.books.open(os.path.abspath(xlsx_path))
            try:
                if sheet_name:
                    try:
                        sheet = wb.sheets[sheet_name]
                    except Exception:
                        raise ValueError(
                            f"Sheet '{sheet_name}' not found. "
                            f"Available: {[s.name for s in wb.sheets]}"
                        )
                    sheets_data = [_serialise_sheet(sheet, include_range=cell_range)]
                else:
                    sheets_data = [_serialise_sheet(s) for s in wb.sheets]

                return {
                    "source_file": os.path.abspath(xlsx_path),
                    "snapshot_version": "1.0",
                    "sheets": sheets_data,
                }
            finally:
                wb.close()
        finally:
            app.quit()

    except Exception as exc:
        if selected_engine == "xlwings":
            raise
        # Auto-fallback to openpyxl headless engine
        return _read_xlsx_openpyxl(xlsx_path, sheet_name=sheet_name, cell_range=cell_range)


def read_xlsx_to_json(
    xlsx_path: str,
    output_path: str | None = None,
    sheet_name: str | None = None,
    cell_range: str | None = None,
    engine: str = "auto",
) -> str:
    """
    Read XLSX and write JSON snapshot to output_path (or return as string).

    Args:
        xlsx_path:   Path to source .xlsx file.
        output_path: Optional path to write JSON output.
        sheet_name:  Optional specific sheet to export.
        cell_range:  Optional cell range (requires sheet_name).
        engine:      'auto', 'openpyxl', or 'xlwings'.

    Returns:
        output_path if written to file, else the JSON string.
    """
    snapshot = read_xlsx(xlsx_path, sheet_name=sheet_name, cell_range=cell_range, engine=engine)
    json_str = json.dumps(snapshot, ensure_ascii=False, indent=2, default=str)

    if output_path:
        out_dir = os.path.dirname(os.path.abspath(output_path))
        if out_dir:
            os.makedirs(out_dir, exist_ok=True)
        with open(output_path, "w", encoding="utf-8") as f:
            f.write(json_str)
        return output_path

    return json_str

