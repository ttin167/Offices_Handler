# -*- coding: utf-8 -*-
"""
xlsx_writer.py — Universal Enterprise Excel (XLSX) Mutation & Generation Engine.

Engine Standards & Invariants:
- ERR_XLSX_001: Prototype Row Style Cloning (100% clone of Font, Fill, Border, Alignment, NumberFormat).
- ERR_XLSX_002: Dynamic Formula Range Expansion & Regex AST Shifting (F10:F25 -> F10:F35, F26 -> F36).
- ERR_XLSX_003: Semantic Anchor Discovery (Keyword header scanning to locate dynamic insertion points).
- ERR_XLSX_004: Safe Merged-Cell Introspection & Border Synchronization across merged ranges.
- ERR_XLSX_005: 2D Freeze Panes & Dynamic Line Auto-Scaling (row height based on wrapped lines).
- ERR_XLSX_006: DrawingML & Package Integrity Preservation (in-place template mutation preserves logos/shapes).
- ERR_XLSX_007: Data Type Coercion & Leading-Zero Loss Guard (Identifier columns strictly set to '@' string format).
- E1-E14: Enterprise UX and template invariants.
"""

from __future__ import annotations

import copy
import json
import os
import re
from typing import Any, Dict, List, Optional, Tuple, Union

import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter, coordinate_to_tuple, column_index_from_string
from openpyxl.worksheet.worksheet import Worksheet

__all__ = [
    "UniversalExcelEngine",
    "mutate_template_excel",
    "safe_set_cell",
    "shift_formula_string",
    "normalize_hex_color",
    "hex_to_rgb_tuple",
    "update_template_workbook",
    "write_xlsx",
    "write_xlsx_from_json_file",
]


def normalize_hex_color(hex_str: str | None) -> str | None:
    """Normalize any hex color format to 6-char RRGGBB or 8-char AARRGGBB."""
    if not hex_str:
        return None
    h = str(hex_str).lstrip("#").strip().upper()
    if len(h) == 8:
        alpha = h[0:2]
        rgb = h[2:8]
        if alpha == "00":
            return f"FF{rgb}"
        return h
    elif len(h) == 6:
        return f"FF{h}"
    return None


def hex_to_rgb_tuple(hex_color: str | None) -> tuple[int, int, int] | None:
    """Convert '#RRGGBB' or 'AARRGGBB' hex string to (R, G, B) tuple."""
    if not hex_color:
        return None
    h = str(hex_color).lstrip("#").strip().upper()
    if len(h) == 8:
        h = h[2:8]
    if len(h) != 6:
        return None
    try:
        return (int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))
    except ValueError:
        return None


_FORMULA_REF_PATTERN = re.compile(
    r"((?:'[^']+'|[A-Za-z0-9_]+)!)?(\$?[A-Za-z]{1,3})(\$?[0-9]+)(?:\s*:\s*(\$?[A-Za-z]{1,3})(\$?[0-9]+))?"
)


def shift_formula_string(
    formula: str,
    insert_row: Optional[int] = None,
    delta: int = 0,
    insert_col: Optional[int] = None,
    delta_cols: int = 0,
    delta_rows: Optional[int] = None,
) -> str:
    """
    Shifts formula cell and range references via Regex AST across rows and columns (ERR_XLSX_002):
    - Row shifting: Ranges spanning across insert_row expand; downstream rows shift down.
    - Cross-sheet references ('Sheet'!$A$1:$B$10) are preserved with proper prefix isolation.
    - Absolute markers ($) on columns and rows are preserved.
    """
    if not formula:
        return formula

    formula_str = str(formula)
    if (
        not formula_str.startswith("=")
        and "!" not in formula_str
        and ":" not in formula_str
        and not re.search(r"\\$?[A-Za-z]{1,3}\\$?[0-9]+", formula_str)
    ):
        return formula

    eff_delta_rows = delta_rows if delta_rows is not None else delta

    def repl(m: re.Match) -> str:
        sheet_prefix = m.group(1) or ""
        c1_s, r1_s = m.group(2), m.group(3)
        c2_s, r2_s = m.group(4), m.group(5)

        c1_name = c1_s.replace("$", "").upper()
        c1_dollar = "$" if "$" in c1_s else ""
        c1_num = column_index_from_string(c1_name)

        r1_num = int(r1_s.replace("$", ""))
        r1_dollar = "$" if "$" in r1_s else ""

        if c2_s and r2_s:
            c2_name = c2_s.replace("$", "").upper()
            c2_dollar = "$" if "$" in c2_s else ""
            c2_num = column_index_from_string(c2_name)

            r2_num = int(r2_s.replace("$", ""))
            r2_dollar = "$" if "$" in r2_s else ""

            # Row shifting
            if insert_row is not None and eff_delta_rows != 0:
                if r1_num <= insert_row and r2_num >= insert_row:
                    r2_num += eff_delta_rows
                elif r1_num > insert_row:
                    r1_num += eff_delta_rows
                    r2_num += eff_delta_rows

            # Column shifting
            if insert_col is not None and delta_cols != 0:
                if c1_num <= insert_col and c2_num >= insert_col:
                    c2_num += delta_cols
                elif c1_num > insert_col:
                    c1_num += delta_cols
                    c2_num += delta_cols

            new_c1 = f"{c1_dollar}{get_column_letter(c1_num)}"
            new_c2 = f"{c2_dollar}{get_column_letter(c2_num)}"
            return f"{sheet_prefix}{new_c1}{r1_dollar}{r1_num}:{new_c2}{r2_dollar}{r2_num}"
        else:
            # Single cell reference
            if insert_row is not None and eff_delta_rows != 0:
                if r1_num >= insert_row:
                    r1_num += eff_delta_rows
            if insert_col is not None and delta_cols != 0:
                if c1_num >= insert_col:
                    c1_num += delta_cols

            new_c1 = f"{c1_dollar}{get_column_letter(c1_num)}"
            return f"{sheet_prefix}{new_c1}{r1_dollar}{r1_num}"

    return _FORMULA_REF_PATTERN.sub(repl, formula)



def safe_set_cell(
    ws: Worksheet,
    coordinate: str,
    value: Any = None,
    formula: str | None = None,
    fill_color: str | None = None,
    font_bold: bool | None = None,
    font_color: str | None = None,
    number_format: str | None = None,
    alignment: Alignment | None = None,
    border: Border | None = None,
) -> Any:
    """
    Safely assign value, formula, and formatting to a cell.
    If the target coordinate is inside a MergedCell, automatically targets
    the Top-Left cell of the range to prevent AttributeError: MergedCell is read-only (ERR_XLSX_004).
    """
    target_coord = coordinate
    target_cell = ws[coordinate]

    if isinstance(target_cell, openpyxl.cell.cell.MergedCell):
        for m_range in ws.merged_cells.ranges:
            if coordinate in m_range:
                target_coord = m_range.coord.split(":")[0]
                break

    cell = ws[target_coord]

    # Value / Formula assignment
    if formula and str(formula).startswith("="):
        cell.value = str(formula)
    elif value is not None:
        cell.value = value

    # Number Format
    if number_format:
        cell.number_format = str(number_format)

    # Shading Fill
    if fill_color:
        norm_fill = normalize_hex_color(fill_color)
        if norm_fill:
            cell.fill = PatternFill(fill_type="solid", start_color=norm_fill, end_color=norm_fill)

    # Font
    if font_bold is not None or font_color is not None:
        curr_font = cell.font or Font()
        norm_fc = normalize_hex_color(font_color) if font_color else (curr_font.color.rgb if curr_font.color else None)
        bold_val = font_bold if font_bold is not None else curr_font.bold
        cell.font = Font(
            name=curr_font.name or "Segoe UI",
            size=curr_font.size or 10,
            bold=bold_val,
            italic=curr_font.italic,
            color=norm_fc
        )

    # Alignment
    if alignment:
        cell.alignment = alignment

    # Border
    if border:
        cell.border = border

    return cell


class UniversalExcelEngine:
    """
    Comprehensive Enterprise OpenXML Excel Engine.
    Implements Dynamic Row Expansion, Prototype Style Cloning, Semantic Discovery,
    and Invariant Enforcement (ERR_XLSX_001 - 007 & E1-E14).
    """

    def __init__(self, workbook_or_path: Union[str, openpyxl.Workbook]):
        if isinstance(workbook_or_path, str):
            if not os.path.isfile(workbook_or_path):
                raise FileNotFoundError(f"Template workbook not found: {workbook_or_path}")
            # Load with data_only=False to preserve DrawingML & formulas (ERR_XLSX_006)
            self.wb: openpyxl.Workbook = openpyxl.load_workbook(workbook_or_path, data_only=False)
            self.source_path: Optional[str] = os.path.abspath(workbook_or_path)
        else:
            self.wb = workbook_or_path
            self.source_path = None

    def get_sheet(self, sheet_name: str) -> Worksheet:
        if sheet_name in self.wb.sheetnames:
            return self.wb[sheet_name]
        return self.wb.create_sheet(title=sheet_name)

    def find_anchor(self, sheet_name: str, keyword: str, max_row: int = 150) -> Optional[Tuple[int, int]]:
        """
        Semantic Anchor Discovery (ERR_XLSX_003):
        Scans sheet matrix to discover the row and column containing `keyword`.
        """
        ws = self.get_sheet(sheet_name)
        kw_clean = str(keyword).strip().lower()

        for r in range(1, min(ws.max_row + 1, max_row + 1)):
            for c in range(1, min(ws.max_column + 1, 40)):
                val = ws.cell(row=r, column=c).value
                if val is not None and kw_clean in str(val).strip().lower():
                    return (r, c)
        return None

    @staticmethod
    def clone_cell_style(source_cell, target_cell, preserve_number_format: bool = True) -> None:
        """
        Prototype Row Style Cloning (ERR_XLSX_001):
        Clones 100% formatting attributes from source_cell onto target_cell.
        """
        if source_cell.has_style:
            if source_cell.font:
                target_cell.font = copy.copy(source_cell.font)
            if source_cell.fill:
                target_cell.fill = copy.copy(source_cell.fill)
            if source_cell.border:
                target_cell.border = copy.copy(source_cell.border)
            if source_cell.alignment:
                target_cell.alignment = copy.copy(source_cell.alignment)
            if preserve_number_format and source_cell.number_format:
                target_cell.number_format = source_cell.number_format

    def expand_table_rows(
        self,
        sheet_name: str,
        start_row: int,
        count: int,
        prototype_row: Optional[int] = None,
        id_columns: Optional[List[int]] = None,
        shift_formulas: bool = True,
    ) -> List[int]:
        """
        Inserts `count` rows at `start_row`, clones prototype formatting,
        and shifts downstream formulas using Regex AST (ERR_XLSX_002).
        Returns the list of newly created row indices.
        """
        if count <= 0:
            return []

        ws = self.get_sheet(sheet_name)
        proto_row_idx = prototype_row if prototype_row is not None else max(1, start_row - 1)
        id_cols = set(id_columns or [])

        # 1. Snapshot prototype row styling BEFORE insert_rows to prevent row-shift index desync (ERR_XLSX_001)
        proto_styles = []
        max_col = max(ws.max_column, 25)
        for c in range(1, max_col + 1):
            proto_cell = ws.cell(row=proto_row_idx, column=c)
            proto_styles.append({
                "font": copy.copy(proto_cell.font) if proto_cell.font else None,
                "fill": copy.copy(proto_cell.fill) if proto_cell.fill else None,
                "border": copy.copy(proto_cell.border) if proto_cell.border else None,
                "alignment": copy.copy(proto_cell.alignment) if proto_cell.alignment else None,
                "number_format": proto_cell.number_format,
            })
        proto_h = ws.row_dimensions[proto_row_idx].height

        # 2. Insert blank rows into openpyxl worksheet
        ws.insert_rows(start_row, count)
        new_row_indices = list(range(start_row, start_row + count))

        # 3. Clone prototype row styling to all newly inserted rows
        for new_r in new_row_indices:
            if proto_h:
                ws.row_dimensions[new_r].height = proto_h

            for c_idx, s in enumerate(proto_styles, start=1):
                target_cell = ws.cell(row=new_r, column=c_idx)
                if s["font"]:
                    target_cell.font = copy.copy(s["font"])
                if s["fill"]:
                    target_cell.fill = copy.copy(s["fill"])
                if s["border"]:
                    target_cell.border = copy.copy(s["border"])
                if s["alignment"]:
                    target_cell.alignment = copy.copy(s["alignment"])
                if s["number_format"]:
                    target_cell.number_format = s["number_format"]

        # 4. Shift merged cell ranges that lie below start_row (ERR_XLSX_004)
        new_merged_ops = []
        for m_range in list(ws.merged_cells.ranges):
            min_col, min_row, max_col, max_row = m_range.bounds
            if min_row >= start_row:
                new_coord = f"{get_column_letter(min_col)}{min_row + count}:{get_column_letter(max_col)}{max_row + count}"
                new_merged_ops.append((m_range, new_coord))
            elif min_row < start_row and max_row >= start_row:
                new_coord = f"{get_column_letter(min_col)}{min_row}:{get_column_letter(max_col)}{max_row + count}"
                new_merged_ops.append((m_range, new_coord))

        for old_range, new_coord in new_merged_ops:
            ws.merged_cells.remove(old_range)
            ws.merge_cells(new_coord)

        # 5. Shift formulas throughout the entire sheet downstream (ERR_XLSX_002)
        if shift_formulas:
            self.shift_formulas_in_sheet(sheet_name, insert_row=start_row, delta_rows=count)

        # 6. Re-anchor and relink charts (Invariant E12)
        self.relink_and_reanchor_charts(sheet_name, insert_row=start_row, delta_rows=count)

        return new_row_indices

    def relink_and_reanchor_charts(
        self,
        sheet_name: str,
        insert_row: Optional[int] = None,
        delta_rows: int = 0,
        insert_col: Optional[int] = None,
        delta_cols: int = 0,
    ) -> None:
        """
        Automated Chart Re-Anchoring & Series Relinking (Invariant E12):
        Shifts chart anchor positions if placed below/right of insertion point,
        and rewrites series formula references (numRef.f, strRef.f) to match new data geometry.
        """
        ws = self.get_sheet(sheet_name)
        charts = getattr(ws, "_charts", [])
        if not charts:
            return

        for chart in charts:
            anchor = getattr(chart, "anchor", None)
            if isinstance(anchor, str):
                m_anch = re.match(r"^([A-Za-z]+)([0-9]+)$", anchor)
                if m_anch:
                    c_letter, r_num = m_anch.group(1), int(m_anch.group(2))
                    c_num = column_index_from_string(c_letter)
                    if insert_row is not None and delta_rows != 0 and r_num >= insert_row:
                        r_num += delta_rows
                    if insert_col is not None and delta_cols != 0 and c_num >= insert_col:
                        c_num += delta_cols
                    chart.anchor = f"{get_column_letter(c_num)}{r_num}"
            elif anchor:
                _from = getattr(anchor, "_from", None)
                to = getattr(anchor, "to", None)
                if _from and hasattr(_from, "row"):
                    if insert_row is not None and delta_rows != 0 and _from.row >= insert_row - 1:
                        _from.row += delta_rows
                    if insert_col is not None and delta_cols != 0 and _from.col >= insert_col - 1:
                        _from.col += delta_cols
                if to and hasattr(to, "row"):
                    if insert_row is not None and delta_rows != 0 and to.row >= insert_row - 1:
                        to.row += delta_rows
                    if insert_col is not None and delta_cols != 0 and to.col >= insert_col - 1:
                        to.col += delta_cols

            for s in getattr(chart, "series", []):
                val = getattr(s, "val", None)
                if val:
                    num_ref = getattr(val, "numRef", None)
                    if num_ref and hasattr(num_ref, "f") and num_ref.f:
                        num_ref.f = shift_formula_string(
                            num_ref.f,
                            insert_row=insert_row,
                            delta_rows=delta_rows,
                            insert_col=insert_col,
                            delta_cols=delta_cols,
                        )
                cat = getattr(s, "cat", None)
                if cat:
                    str_ref = getattr(cat, "strRef", None)
                    if str_ref and hasattr(str_ref, "f") and str_ref.f:
                        str_ref.f = shift_formula_string(
                            str_ref.f,
                            insert_row=insert_row,
                            delta_rows=delta_rows,
                            insert_col=insert_col,
                            delta_cols=delta_cols,
                        )
                title = getattr(s, "title", None)
                if title:
                    str_ref = getattr(title, "strRef", None)
                    if str_ref and hasattr(str_ref, "f") and str_ref.f:
                        str_ref.f = shift_formula_string(
                            str_ref.f,
                            insert_row=insert_row,
                            delta_rows=delta_rows,
                            insert_col=insert_col,
                            delta_cols=delta_cols,
                        )

    def shift_formulas_in_sheet(
        self,
        sheet_name: str,
        insert_row: Optional[int] = None,
        delta_rows: int = 0,
        insert_col: Optional[int] = None,
        delta_cols: int = 0,
    ) -> None:
        """Scans and updates all formula cells in sheet using multi-axis shift_formula_string."""
        ws = self.get_sheet(sheet_name)
        for r in range(1, ws.max_row + 1):
            for c in range(1, ws.max_column + 1):
                cell = ws.cell(row=r, column=c)
                val = cell.value
                if isinstance(val, str) and val.startswith("="):
                    cell.value = shift_formula_string(
                        val,
                        insert_row=insert_row,
                        delta_rows=delta_rows,
                        insert_col=insert_col,
                        delta_cols=delta_cols,
                    )

    def expand_table_columns(
        self,
        sheet_name: str,
        start_col: int,
        count: int,
        prototype_col: Optional[int] = None,
        min_row: int = 1,
        max_row: Optional[int] = None,
        shift_formulas: bool = True,
    ) -> List[int]:
        """
        Inserts `count` columns at `start_col`, clones prototype column geometry & styling
        (including header rotation, cell borders, and fills), shifts downstream merged cells
        and formulas, and automatically re-anchors charts (ERR_XLSX_001, Invariant E5 & E12).
        """
        if count <= 0:
            return []

        ws = self.get_sheet(sheet_name)
        proto_c_idx = prototype_col if prototype_col is not None else max(1, start_col - 1)
        proto_letter = get_column_letter(proto_c_idx)
        proto_width = ws.column_dimensions[proto_letter].width

        # 1. Snapshot prototype column cell styles across rows
        end_row = max_row if max_row is not None else max(ws.max_row, 30)
        proto_styles = []
        for r in range(min_row, end_row + 1):
            p_cell = ws.cell(row=r, column=proto_c_idx)
            proto_styles.append({
                "row": r,
                "font": copy.copy(p_cell.font) if p_cell.font else None,
                "fill": copy.copy(p_cell.fill) if p_cell.fill else None,
                "border": copy.copy(p_cell.border) if p_cell.border else None,
                "alignment": copy.copy(p_cell.alignment) if p_cell.alignment else None,
                "number_format": p_cell.number_format,
            })

        # 2. Insert blank columns into openpyxl worksheet
        ws.insert_cols(start_col, count)
        new_col_indices = list(range(start_col, start_col + count))

        # 3. Apply prototype geometry and styling to all newly inserted columns
        for new_c in new_col_indices:
            new_letter = get_column_letter(new_c)
            if proto_width:
                ws.column_dimensions[new_letter].width = proto_width

            for s in proto_styles:
                target_cell = ws.cell(row=s["row"], column=new_c)
                if s["font"]:
                    target_cell.font = copy.copy(s["font"])
                if s["fill"]:
                    target_cell.fill = copy.copy(s["fill"])
                if s["border"]:
                    target_cell.border = copy.copy(s["border"])
                if s["alignment"]:
                    target_cell.alignment = copy.copy(s["alignment"])
                if s["number_format"]:
                    target_cell.number_format = s["number_format"]

        # 4. Shift merged cell ranges that span across or lie to the right of start_col
        new_merged_ops = []
        for m_range in list(ws.merged_cells.ranges):
            min_c, min_r, max_c, max_r = m_range.bounds
            if min_c >= start_col:
                new_coord = f"{get_column_letter(min_c + count)}{min_r}:{get_column_letter(max_c + count)}{max_r}"
                new_merged_ops.append((m_range, new_coord))
            elif min_c < start_col and max_c >= start_col:
                new_coord = f"{get_column_letter(min_c)}{min_r}:{get_column_letter(max_c + count)}{max_r}"
                new_merged_ops.append((m_range, new_coord))

        for old_range, new_coord in new_merged_ops:
            ws.merged_cells.remove(old_range)
            ws.merge_cells(new_coord)

        # 5. Shift formulas horizontally across the entire sheet
        if shift_formulas:
            self.shift_formulas_in_sheet(sheet_name, insert_col=start_col, delta_cols=count)

        # 6. Re-anchor and relink charts
        self.relink_and_reanchor_charts(sheet_name, insert_col=start_col, delta_cols=count)

        return new_col_indices

    def populate_table_data(
        self,
        sheet_name: str,
        start_row: int,
        rows_data: List[List[Any]],
        prototype_row: Optional[int] = None,
        id_columns: Optional[List[int]] = None,
        start_col: int = 1,
    ) -> None:
        """
        Populates a block of rows into the sheet, ensuring values and formulas are written
        and enforcing identifier types and string number formatting (ERR_XLSX_007).
        """
        ws = self.get_sheet(sheet_name)
        id_cols = set(id_columns or [])

        for r_offset, row_vals in enumerate(rows_data):
            curr_row = start_row + r_offset
            for c_offset, val in enumerate(row_vals):
                curr_col = start_col + c_offset
                cell = ws.cell(row=curr_row, column=curr_col)

                is_id = (c_offset in id_cols) or (curr_col in id_cols)

                if is_id:
                    cell.number_format = "@"
                    cell.value = str(val) if val is not None else ""
                elif isinstance(val, str) and val.startswith("="):
                    cell.value = val
                else:
                    cell.value = val

    def auto_fit_layout(
        self,
        sheet_name: str,
        min_col_width: float = 10.0,
        max_col_width: float = 65.0,
        col_padding: float = 3.0,
        line_height: float = 14.5,
        min_row_height: float = 20.0,
    ) -> None:
        """
        UX Dynamic Text Auto-Scaling (E13 & ERR_XLSX_005):
        Adjusts column widths based on maximum content length and row heights
        based on multi-line text wrapping to eliminate truncated text.
        """
        ws = self.get_sheet(sheet_name)

        # 1. Adjust column widths
        for col in ws.columns:
            first_cell = col[0]
            col_letter = get_column_letter(first_cell.column)
            max_len = 0

            for cell in col:
                val = cell.value
                if val is not None:
                    # Ignore formulas for width calculation
                    val_str = str(val)
                    if not val_str.startswith("="):
                        lines = val_str.split("\n")
                        max_len = max(max_len, max(len(line) for line in lines))

            if max_len > 0:
                current_dim = ws.column_dimensions[col_letter].width or min_col_width
                calculated_w = max(min_col_width, min(max_col_width, max_len + col_padding))
                ws.column_dimensions[col_letter].width = max(current_dim, calculated_w)

        # 2. Adjust row heights for wrapped text (ERR_XLSX_004 / Invariant E13)
        for r in range(1, ws.max_row + 1):
            max_lines = 1
            for c in range(1, ws.max_column + 1):
                cell = ws.cell(row=r, column=c)
                val = cell.value
                if val is not None and isinstance(val, str) and not val.startswith("="):
                    val_str = str(val)
                    split_lines = len(val_str.split("\n"))
                    wrap_lines = int(len(val_str) / 45.0)
                    total_lines = split_lines + wrap_lines
                    if total_lines > max_lines:
                        max_lines = total_lines

                    # Automatically set wrap_text for long text (ERR_XLSX_004 / E13)
                    if len(val_str) > 40:
                        cur_al = cell.alignment
                        if not cur_al or not cur_al.wrap_text:
                            cell.alignment = Alignment(
                                horizontal=cur_al.horizontal if cur_al else "left",
                                vertical=cur_al.vertical if cur_al else "center",
                                wrap_text=True
                            )

            if max_lines > 1:
                current_h = ws.row_dimensions[r].height or min_row_height
                needed_h = max(min_row_height, max_lines * 14.5)
                ws.row_dimensions[r].height = max(current_h, needed_h)

    def sync_merged_borders(self, sheet_name: str) -> None:
        """
        Safe Merged-Cell Introspection & Border Synchronization (ERR_XLSX_004):
        Ensures all cells within every merged range share consistent outer borders
        to prevent jagged or missing borders.
        """
        ws = self.get_sheet(sheet_name)
        for m_range in list(ws.merged_cells.ranges):
            min_col, min_row, max_col, max_row = m_range.bounds
            top_left = ws.cell(row=min_row, column=min_col)
            border_template = top_left.border

            if not border_template:
                continue

            for r in range(min_row, max_row + 1):
                for c in range(min_col, max_col + 1):
                    cell = ws.cell(row=r, column=c)
                    cell.border = Border(
                        top=border_template.top if r == min_row else Side(style=None),
                        bottom=border_template.bottom if r == max_row else Side(style=None),
                        left=border_template.left if c == min_col else Side(style=None),
                        right=border_template.right if c == max_col else Side(style=None),
                    )

    def set_freeze_panes(self, sheet_name: str, coordinate: str = "B8") -> None:
        """Applies 2D Freeze Panes at specified coordinate (ERR_XLSX_005)."""
        ws = self.get_sheet(sheet_name)
        ws.freeze_panes = coordinate

    def save(self, output_path: str) -> str:
        """Saves mutated workbook preserving DrawingML packages (ERR_XLSX_006)."""
        dest = os.path.abspath(output_path)
        out_dir = os.path.dirname(dest)
        if out_dir:
            os.makedirs(out_dir, exist_ok=True)
        self.wb.save(dest)
        return dest


def mutate_template_excel(
    template_path: str,
    output_path: str,
    spec: dict[str, Any],
) -> str:
    """
    High-level declarative mutation of an Excel template workbook.

    Spec format:
    {
      "sheets": {
        "SheetName": {
          "cell_updates": {
            "A4": "Project Name",
            "B6": {"value": 15, "font_bold": True},
            "E16": {"formula": "=SUM(E10:E15)"}
          },
          "table_expansions": [
            {
              "anchor_keyword": "Test Case ID",   // or start_row: 10
              "prototype_row": 10,
              "rows": [
                ["TC_001", "Verify Login", "Success", "=IF(...)"],
                ["TC_002", "Verify Checkout", "Success", "=IF(...)"]
              ],
              "id_columns": [0],
              "auto_fit": true
            }
          ],
          "freeze_panes": "B8"
        }
      }
    }
    """
    engine = UniversalExcelEngine(template_path)
    sheets_spec = spec.get("sheets", {})

    for sheet_name, s_spec in sheets_spec.items():
        ws = engine.get_sheet(sheet_name)

        # 1. Table Expansions (Dynamic row insertion & style cloning)
        expansions = s_spec.get("table_expansions", [])
        for exp in expansions:
            rows_data = exp.get("rows", [])
            if not rows_data:
                continue

            # Determine start_row via anchor or explicit number
            start_row = exp.get("start_row")
            anchor_kw = exp.get("anchor_keyword")
            if not start_row and anchor_kw:
                pos = engine.find_anchor(sheet_name, anchor_kw)
                if pos:
                    start_row = pos[0] + 1  # Row right below header sentinel

            if not start_row:
                start_row = 10  # Standard fallback data row

            proto_row = exp.get("prototype_row", start_row)
            id_cols = exp.get("id_columns", [0])

            # Expand rows and populate
            engine.expand_table_rows(
                sheet_name=sheet_name,
                start_row=start_row,
                count=len(rows_data),
                prototype_row=proto_row,
                id_columns=id_cols,
                shift_formulas=True
            )
            engine.populate_table_data(
                sheet_name=sheet_name,
                start_row=start_row,
                rows_data=rows_data,
                prototype_row=proto_row,
                id_columns=id_cols
            )

            if exp.get("auto_fit", True):
                engine.auto_fit_layout(sheet_name)

        # 1b. Column Expansions (Dynamic column insertion & prototype styling)
        col_expansions = s_spec.get("column_expansions", [])
        for exp in col_expansions:
            start_col = exp.get("start_col")
            anchor_kw = exp.get("anchor_keyword")
            if not start_col and anchor_kw:
                pos = engine.find_anchor(sheet_name, anchor_kw)
                if pos:
                    start_col = pos[1] + 1

            if not start_col:
                start_col = 6  # Standard fallback column (Col F)

            count = exp.get("count", 1)
            proto_col = exp.get("prototype_col", start_col)
            min_r = exp.get("min_row", 1)
            max_r = exp.get("max_row", None)

            engine.expand_table_columns(
                sheet_name=sheet_name,
                start_col=start_col,
                count=count,
                prototype_col=proto_col,
                min_row=min_r,
                max_row=max_r,
                shift_formulas=True,
            )

        # 2. Cell Updates
        cell_updates = s_spec.get("cell_updates", {})
        for coord, cell_info in cell_updates.items():
            if isinstance(cell_info, dict):
                safe_set_cell(
                    ws=ws,
                    coordinate=coord,
                    value=cell_info.get("value"),
                    formula=cell_info.get("formula"),
                    fill_color=cell_info.get("fill_color"),
                    font_bold=cell_info.get("font_bold"),
                    font_color=cell_info.get("font_color"),
                    number_format=cell_info.get("number_format")
                )
            else:
                safe_set_cell(ws=ws, coordinate=coord, value=cell_info)

        # 3. Freeze Panes
        fp = s_spec.get("freeze_panes")
        if fp:
            engine.set_freeze_panes(sheet_name, fp)

        # 4. Merged Borders Sync
        engine.sync_merged_borders(sheet_name)

    return engine.save(output_path)


# ---------------------------------------------------------------------------
# Backward Compatible APIs
# ---------------------------------------------------------------------------

def update_template_workbook(
    template_path: str,
    output_path: str,
    sheet_updates: dict[str, dict[str, Any]]
) -> str:
    """Backward-compatible template cell updater."""
    spec = {
        "sheets": {
            sname: {"cell_updates": updates} for sname, updates in sheet_updates.items()
        }
    }
    return mutate_template_excel(template_path, output_path, spec)


def write_xlsx(
    snapshot: dict[str, Any],
    output_path: str,
    template_path: str | None = None,
) -> str:
    """Write an XLSX file from a JSON snapshot."""
    out_dir = os.path.dirname(os.path.abspath(output_path))
    if out_dir:
        os.makedirs(out_dir, exist_ok=True)

    if template_path and os.path.isfile(template_path):
        wb = openpyxl.load_workbook(template_path, data_only=False)
    else:
        wb = openpyxl.Workbook()
        if "Sheet" in wb.sheetnames and len(snapshot.get("sheets", [])) > 0:
            wb.remove(wb["Sheet"])

    for sheet_data in snapshot.get("sheets", []):
        sname = sheet_data.get("sheet_name", "Sheet1")
        if sname in wb.sheetnames:
            ws = wb[sname]
        else:
            ws = wb.create_sheet(title=sname)

        for cell_data in sheet_data.get("cells", []):
            addr = cell_data.get("address")
            if not addr:
                continue
            fmt = cell_data.get("format", {})
            safe_set_cell(
                ws=ws,
                coordinate=addr,
                value=cell_data.get("value"),
                formula=cell_data.get("formula"),
                fill_color=fmt.get("fill_color"),
                font_bold=fmt.get("bold"),
                font_color=fmt.get("font_color")
            )

    wb.save(output_path)
    return os.path.abspath(output_path)


def write_xlsx_from_json_file(
    json_path: str,
    output_path: str,
    template_path: str | None = None,
) -> str:
    """Load JSON snapshot and write to XLSX."""
    if not os.path.isfile(json_path):
        raise FileNotFoundError(f"JSON snapshot not found: {json_path}")

    with open(json_path, "r", encoding="utf-8") as f:
        snapshot = json.load(f)

    return write_xlsx(snapshot, output_path, template_path)
