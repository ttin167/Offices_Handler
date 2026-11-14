"""
core.xlsx.power_tools — Pure Office Excel Power Suite.

Headless OpenPyXL automation for everyday Office IT workflows:
1. VN Data Cleaner & Standardizer (dedup, whitespace, date/phone/CCCD normalization).
2. Split / Merge Columns (including Vietnamese Full Name splitting).
3. Pivot Summary + Embedded Native Chart (live SUMIFS/COUNTIFS/AVERAGEIFS formulas).
4. No-Code XLOOKUP Join across sheets or workbooks.
5. Split Master Workbook by Column Value into individual XLSX files + ZIP archive.
6. Visual Excel Diff (cell-by-cell highlighting + Comment annotations + _Diff_Summary).

Follows OpenXML invariants ERR_XLSX_001..007 & E1..E14 strictly without COM dependencies.
"""

from __future__ import annotations

import datetime
import os
import re
import tempfile
import zipfile
from copy import copy
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence, Tuple, Union

import openpyxl
from openpyxl.chart import BarChart, LineChart, PieChart, Reference
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

from .writer import (
    apply_worksheet_print_setup,
    parse_vietnamese_date,
    sanitize_xml_string,
)

# Shared palette tokens (accessible slate/navy enterprise styles)
_HEADER_FILL = PatternFill("solid", fgColor="1E3A5F")
_HEADER_FONT = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
_SUBTOTAL_FILL = PatternFill("solid", fgColor="E2E8F0")
_SUBTOTAL_FONT = Font(name="Calibri", size=11, bold=True, color="0F172A")
_ZEBRA_FILL = PatternFill("solid", fgColor="F8FAFC")
_THIN_SIDE = Side(style="thin", color="CBD5E1")
_MED_NAVY_SIDE = Side(style="medium", color="1E3A5F")
_CELL_BORDER = Border(left=_THIN_SIDE, right=_THIN_SIDE, top=_THIN_SIDE, bottom=_THIN_SIDE)
_HEADER_BORDER = Border(
    left=_THIN_SIDE, right=_THIN_SIDE, top=_MED_NAVY_SIDE, bottom=_MED_NAVY_SIDE
)

_DIFF_MODIFIED_FILL = PatternFill("solid", fgColor="FEF3C7")
_DIFF_MODIFIED_FONT = Font(name="Calibri", size=11, bold=False, color="92400E")
_DIFF_ADDED_FILL = PatternFill("solid", fgColor="D1FAE5")
_DIFF_ADDED_FONT = Font(name="Calibri", size=11, bold=False, color="065F46")
_DIFF_REMOVED_FILL = PatternFill("solid", fgColor="FEE2E2")
_DIFF_REMOVED_FONT = Font(name="Calibri", size=11, bold=False, color="991B1B")

_ID_PHONE_KEYWORDS = (
    "sdt",
    "sđt",
    "phone",
    "mobile",
    "tel",
    "điện thoại",
    "dien thoai",
    "dienthoai",
    "di động",
    "didong",
    "cccd",
    "cmnd",
    "passport",
    "hộ chiếu",
    "ho chieu",
    "stk",
    "tài khoản",
    "tai khoan",
    "mst",
    "mã số thuế",
    "ma so thue",
    "code",
    "mã ",
    "ma_",
    "_id",
)

_PHONE_KEYWORDS = ("sdt", "sđt", "phone", "mobile", "tel", "điện thoại", "dien thoai", "dienthoai", "di động", "didong")


# ---------------------------------------------------------------------------
# Internal Shared Helpers (Guard Clauses, Column Resolution, Style Cloning)
# ---------------------------------------------------------------------------


def _ensure_existing_file(file_path: str, label: str = "Input workbook") -> Path:
    """Validate that a file path exists and return it as a Path."""
    path = Path(file_path)
    if not path.is_file():
        raise FileNotFoundError(f"{label} not found: {file_path}")
    return path


def _resolve_output_path(input_path: Path, output_path: Optional[str], suffix: str) -> Path:
    """Resolve destination path cleanly without overwriting the original unless specified."""
    if output_path:
        out = Path(output_path)
    else:
        out = input_path.with_name(f"{input_path.stem}_{suffix}{input_path.suffix}")
    out.parent.mkdir(parents=True, exist_ok=True)
    return out


def _select_worksheet(wb: openpyxl.Workbook, sheet_name: Optional[str]) -> Worksheet:
    """Select worksheet by exact or stripped name, defaulting to active sheet."""
    if not sheet_name:
        ws = wb.active
        if ws is None:
            raise ValueError("Workbook contains no worksheets.")
        return ws
    if sheet_name in wb.sheetnames:
        return wb[sheet_name]
    lower_map = {name.strip().lower(): name for name in wb.sheetnames}
    key = sheet_name.strip().lower()
    if key in lower_map:
        return wb[lower_map[key]]
    raise ValueError(f"Worksheet '{sheet_name}' not found. Available: {wb.sheetnames}")


def _detect_header_row(ws: Worksheet, fallback: int = 1) -> int:
    """Locate the first non-empty row in a worksheet to serve as header row."""
    max_scan = min(ws.max_row or 1, 15)
    for r in range(1, max_scan + 1):
        row_vals = [ws.cell(r, c).value for c in range(1, (ws.max_column or 1) + 1)]
        if any(v is not None and str(v).strip() != "" for v in row_vals):
            return r
    return fallback


def _get_header_map(ws: Worksheet, header_row: int) -> Tuple[List[str], Dict[str, int]]:
    """Build ordered header list and lookup dict supporting both header names and col letters."""
    max_col = ws.max_column or 0
    headers: List[str] = []
    lookup: Dict[str, int] = {}
    for c in range(1, max_col + 1):
        raw = ws.cell(header_row, c).value
        name = str(raw).strip() if raw is not None and str(raw).strip() else get_column_letter(c)
        headers.append(name)
        lookup.setdefault(name.lower(), c)
        lookup.setdefault(get_column_letter(c).lower(), c)
    return headers, lookup


def _resolve_column_index(col_spec: str, ws: Worksheet, header_row: int) -> int:
    """Resolve a column header name, 1-based letter ('B'), or numeric string into 1-based index."""
    if not col_spec or not str(col_spec).strip():
        raise ValueError("Column specifier cannot be empty.")
    spec = str(col_spec).strip()
    _, lookup = _get_header_map(ws, header_row)
    if spec.lower() in lookup:
        return lookup[spec.lower()]
    if spec.isdigit():
        idx = int(spec)
        if 1 <= idx <= max(ws.max_column or 1, idx):
            return idx
    if re.fullmatch(r"[A-Za-z]{1,3}", spec):
        return column_index_from_string(spec.upper())
    raise ValueError(f"Column '{col_spec}' not found in sheet '{ws.title}'.")


def _clone_cell_style(src_cell: Any, dst_cell: Any) -> None:
    """Clone font, fill, border, alignment, and number_format (ERR_XLSX_001)."""
    if src_cell.has_style:
        dst_cell.font = copy(src_cell.font)
        dst_cell.fill = copy(src_cell.fill)
        dst_cell.border = copy(src_cell.border)
        dst_cell.alignment = copy(src_cell.alignment)
        dst_cell.number_format = src_cell.number_format


def _is_id_or_phone_header(header_name: str) -> Tuple[bool, bool]:
    """Return (is_id_or_phone, is_specifically_phone) based on header keywords."""
    low = header_name.strip().lower()
    is_phone = any(kw in low for kw in _PHONE_KEYWORDS)
    is_id = is_phone or any(kw in low for kw in _ID_PHONE_KEYWORDS) or low == "id"
    return is_id, is_phone


def _normalize_identifier_value(raw: Any, is_phone: bool) -> str:
    """Normalize CCCD/Phone/Code identifiers into clean text preserving leading zeros."""
    if raw is None:
        return ""
    if isinstance(raw, float) and raw.is_integer():
        text = str(int(raw))
    elif isinstance(raw, int):
        text = str(raw)
    else:
        text = str(raw).strip()

    if not text:
        return ""

    # Strip whitespace/dots/hyphens in numeric phone values
    compact = re.sub(r"[\s.\-()]", "", text)
    if is_phone:
        if compact.startswith("+84"):
            compact = "0" + compact[3:]
        elif compact.startswith("84") and len(compact) >= 11:
            compact = "0" + compact[2:]
        elif re.fullmatch(r"(3|5|7|8|9)\d{8}", compact):
            compact = "0" + compact
        return compact

    # General ID / CCCD / Tax code: remove Excel accidental '.0' suffix
    if re.fullmatch(r"\d+\.0+", text):
        return text.split(".", 1)[0]
    return text


def _normalize_key_for_join(val: Any) -> str:
    """Normalize join key across sheets/files so '101', 101, and 101.0 match reliably."""
    if val is None:
        return ""
    if isinstance(val, float) and val.is_integer():
        return str(int(val))
    if isinstance(val, (datetime.datetime, datetime.date)):
        return val.strftime("%Y-%m-%d")
    text = str(val).strip()
    if re.fullmatch(r"\d+\.0+", text):
        return text.split(".", 1)[0]
    return text


def _apply_case_mode(text: str, case_mode: str) -> str:
    """Apply configured case standardization to string content."""
    mode = (case_mode or "preserve").strip().lower()
    if mode == "upper":
        return text.upper()
    if mode == "lower":
        return text.lower()
    if mode in {"title", "vn_name"}:
        return " ".join(w.capitalize() if not w.isupper() or len(w) <= 1 else w for w in text.split())
    return text


def _looks_like_date_string(text: str) -> bool:
    """Guard against coercing plain numeric strings or IDs into dates."""
    if not text or len(text) > 32:
        return False
    return bool(re.search(r"[/\-]|ngày\s+\d", text, re.IGNORECASE)) and bool(re.search(r"\d", text))


def _auto_fit_worksheet_columns(
    ws: Worksheet,
    header_row: int = 1,
    min_width: float = 10.0,
    max_width: float = 48.0,
) -> None:
    """Measure visible text per column and apply cushioned widths + dynamic row heights (E13)."""
    max_col = ws.max_column or 1
    max_row = ws.max_row or 1
    col_widths: Dict[int, float] = {}

    for c in range(1, max_col + 1):
        best = min_width
        for r in range(header_row, max_row + 1):
            val = ws.cell(r, c).value
            if val is None or (isinstance(val, str) and val.startswith("=")):
                continue
            longest_line = max((len(line) for line in str(val).splitlines()), default=0)
            best = max(best, min(max_width, longest_line + 4.0))
        col_widths[c] = best
        ws.column_dimensions[get_column_letter(c)].width = best

    for r in range(header_row, max_row + 1):
        max_lines = 1
        for c in range(1, max_col + 1):
            cell = ws.cell(r, c)
            val = cell.value
            if val is None or (isinstance(val, str) and val.startswith("=")):
                continue
            s = str(val)
            eff_w = max(8.0, col_widths[c] - 2.0)
            wrapped_lines = sum(
                max(1, (len(line) + int(eff_w) - 1) // int(eff_w))
                for line in (s.splitlines() or [""])
            )
            max_lines = max(max_lines, wrapped_lines)
        base_h = 26.0 if r == header_row else 21.0
        ws.row_dimensions[r].height = max(base_h, min(140.0, max_lines * 16.0 + 5.0))


def _split_vietnamese_fullname(full_name: str, num_parts: int = 3) -> List[str]:
    """
    Split a Vietnamese full name cleanly into 3 parts [Họ, Tên đệm, Tên]
    or 2 parts [Họ và tên đệm, Tên].
    """
    cleaned = " ".join(str(full_name or "").split())
    if not cleaned:
        return ["", ""] if num_parts == 2 else ["", "", ""]
    tokens = cleaned.split(" ")
    if num_parts == 2:
        if len(tokens) == 1:
            return ["", tokens[0]]
        return [" ".join(tokens[:-1]), tokens[-1]]
    if len(tokens) == 1:
        return ["", "", tokens[0]]
    if len(tokens) == 2:
        return [tokens[0], "", tokens[1]]
    return [tokens[0], " ".join(tokens[1:-1]), tokens[-1]]


def _safe_filename_slug(name: str, fallback: str = "Sheet") -> str:
    """Create a Windows/POSIX-safe filename component avoiding zip slip."""
    cleaned = re.sub(r'[\\/*?:"<>|\x00-\x1f]+', "_", str(name or "").strip())
    cleaned = cleaned.strip(". ")
    return cleaned[:80] or fallback


def _delete_rows_grouped(ws: Worksheet, row_indices: List[int]) -> int:
    """Delete 1-based row indices from bottom up by grouping contiguous ranges."""
    if not row_indices:
        return 0
    sorted_desc = sorted(set(row_indices), reverse=True)
    deleted = 0
    run_end = sorted_desc[0]
    run_start = sorted_desc[0]

    for idx in sorted_desc[1:]:
        if idx == run_start - 1:
            run_start = idx
        else:
            count = run_end - run_start + 1
            ws.delete_rows(run_start, count)
            deleted += count
            run_end = idx
            run_start = idx

    count = run_end - run_start + 1
    ws.delete_rows(run_start, count)
    deleted += count
    return deleted


# ---------------------------------------------------------------------------
# 1. VN Data Cleaner & Standardizer
# ---------------------------------------------------------------------------


def clean_and_standardize_workbook(
    input_path: str,
    output_path: Optional[str] = None,
    sheet_name: Optional[str] = None,
    dedup_columns: Optional[List[str]] = None,
    normalize_dates: bool = True,
    normalize_phones_and_ids: bool = True,
    trim_whitespace: bool = True,
    case_mode: str = "preserve",
) -> Dict[str, Any]:
    """
    Standardize a messy Excel sheet in-place on a cloned copy of the workbook:
    - Collapse redundant spaces and strip invalid XML control chars.
    - Convert Vietnamese date strings (dd/mm/yyyy, yyyy-mm-dd, ...) to real Excel dates.
    - Protect Phones, CCCD, IDs, Tax Codes as '@' text with leading zeros restored.
    - Remove duplicate rows keyed on `dedup_columns` (or `["*"]` for whole-row dedup).
    """
    src = _ensure_existing_file(input_path)
    dst = _resolve_output_path(src, output_path, "cleaned")

    wb = openpyxl.load_workbook(str(src))
    ws = _select_worksheet(wb, sheet_name)
    header_row = _detect_header_row(ws)
    headers, _ = _get_header_map(ws, header_row)
    max_col = max(ws.max_column or 1, len(headers))
    max_row = ws.max_row or header_row

    # Resolve column categories once
    col_flags: Dict[int, Tuple[bool, bool]] = {
        c: _is_id_or_phone_header(headers[c - 1] if c - 1 < len(headers) else "")
        for c in range(1, max_col + 1)
    }

    # Resolve dedup column indices
    dedup_indices: Optional[List[int]] = None
    if dedup_columns:
        if len(dedup_columns) == 1 and dedup_columns[0].strip() == "*":
            dedup_indices = list(range(1, max_col + 1))
        else:
            dedup_indices = [_resolve_column_index(col, ws, header_row) for col in dedup_columns]

    cells_cleaned = 0
    dates_standardized = 0
    ids_protected = 0
    rows_processed = 0
    seen_keys: set[Tuple[str, ...]] = set()
    dup_rows: List[int] = []

    for r in range(header_row + 1, max_row + 1):
        row_values = [ws.cell(r, c).value for c in range(1, max_col + 1)]
        if all(v is None or str(v).strip() == "" for v in row_values):
            continue
        rows_processed += 1

        for c in range(1, max_col + 1):
            cell = ws.cell(r, c)
            val = cell.value
            if val is None or cell.data_type == "f":
                continue

            is_id, is_phone = col_flags[c]
            if normalize_phones_and_ids and is_id:
                norm_id = _normalize_identifier_value(val, is_phone=is_phone)
                if norm_id != str(val) or cell.number_format != "@":
                    ids_protected += 1
                cell.value = sanitize_xml_string(norm_id)
                cell.data_type = "s"
                cell.number_format = "@"
                continue

            if isinstance(val, str):
                cleaned_str = sanitize_xml_string(val)
                if trim_whitespace:
                    collapsed = " ".join(cleaned_str.split())
                else:
                    collapsed = cleaned_str

                if normalize_dates and _looks_like_date_string(collapsed):
                    parsed_date = parse_vietnamese_date(collapsed)
                    if parsed_date is not None:
                        cell.value = parsed_date
                        cell.number_format = "dd/mm/yyyy"
                        dates_standardized += 1
                        continue

                if case_mode != "preserve" and collapsed:
                    collapsed = _apply_case_mode(collapsed, case_mode)

                if collapsed != val:
                    cells_cleaned += 1
                cell.value = collapsed

        if dedup_indices is not None:
            key = tuple(
                _normalize_key_for_join(ws.cell(r, c).value).lower() for c in dedup_indices
            )
            if any(key):
                if key in seen_keys:
                    dup_rows.append(r)
                else:
                    seen_keys.add(key)

    duplicates_removed = _delete_rows_grouped(ws, dup_rows)
    ws.freeze_panes = ws.cell(header_row + 1, 1).coordinate
    apply_worksheet_print_setup(ws, print_title_header_row=header_row)
    wb.save(str(dst))

    return {
        "ok": True,
        "input_path": str(src),
        "output_path": str(dst),
        "sheet_name": ws.title,
        "header_row": header_row,
        "rows_processed": rows_processed,
        "cells_cleaned": cells_cleaned,
        "dates_standardized": dates_standardized,
        "ids_protected": ids_protected,
        "duplicates_removed": duplicates_removed,
    }


# ---------------------------------------------------------------------------
# 2. Split / Merge Columns (including VN Full Name Splitter)
# ---------------------------------------------------------------------------


def transform_workbook_columns(
    input_path: str,
    mode: str,
    source_columns: Sequence[str],
    output_path: Optional[str] = None,
    sheet_name: Optional[str] = None,
    delimiter: str = " ",
    joiner: str = " ",
    target_columns: Optional[Sequence[str]] = None,
    new_column_name: Optional[str] = None,
) -> Dict[str, Any]:
    """
    Transform columns in an Excel worksheet while keeping existing sheets/styles:
    - `split_vn_fullname`: Split a Vietnamese name column into 3 cols [Họ, Tên đệm, Tên]
      (or 2 cols if `target_columns` provides 2 names).
    - `split_delimiter`: Split one source column by `delimiter`.
    - `merge_columns`: Merge multiple source columns with `joiner` into `new_column_name`.
    """
    norm_mode = (mode or "").strip().lower()
    if norm_mode not in {"split_vn_fullname", "split_delimiter", "merge_columns"}:
        raise ValueError(
            "Unsupported mode. Expected 'split_vn_fullname', 'split_delimiter', or 'merge_columns'."
        )
    if not source_columns:
        raise ValueError("source_columns must contain at least one column.")

    src = _ensure_existing_file(input_path)
    dst = _resolve_output_path(src, output_path, f"col_{norm_mode}")

    wb = openpyxl.load_workbook(str(src))
    ws = _select_worksheet(wb, sheet_name)
    header_row = _detect_header_row(ws)
    max_row = ws.max_row or header_row
    max_col = ws.max_column or 1

    src_indices = [_resolve_column_index(col, ws, header_row) for col in source_columns]
    proto_header = ws.cell(header_row, src_indices[0])
    proto_data = ws.cell(min(header_row + 1, max_row), src_indices[0])

    created_columns: List[str] = []
    rows_transformed = 0

    if norm_mode == "merge_columns":
        col_title = (new_column_name or "Merged_Column").strip()
        target_col_idx = max_col + 1
        hdr_cell = ws.cell(header_row, target_col_idx, value=sanitize_xml_string(col_title))
        _clone_cell_style(proto_header, hdr_cell)
        created_columns.append(col_title)

        for r in range(header_row + 1, max_row + 1):
            parts = []
            for idx in src_indices:
                v = ws.cell(r, idx).value
                if v is not None and str(v).strip() != "":
                    parts.append(str(v).strip())
            if not parts:
                continue
            merged_val = sanitize_xml_string(joiner.join(parts))
            dst_cell = ws.cell(r, target_col_idx, value=merged_val)
            _clone_cell_style(proto_data, dst_cell)
            dst_cell.number_format = "@" if merged_val.startswith("0") and merged_val.isdigit() else "General"
            rows_transformed += 1

    else:
        src_col_idx = src_indices[0]
        src_header_val = str(ws.cell(header_row, src_col_idx).value or "Col").strip()

        # Determine target headers and per-row split arrays
        row_splits: Dict[int, List[str]] = {}
        max_parts = 1
        for r in range(header_row + 1, max_row + 1):
            raw = ws.cell(r, src_col_idx).value
            if raw is None or str(raw).strip() == "":
                continue
            text = str(raw).strip()
            if norm_mode == "split_vn_fullname":
                n_parts = 2 if (target_columns and len(target_columns) == 2) else 3
                parts = _split_vietnamese_fullname(text, num_parts=n_parts)
            else:
                sep = delimiter if delimiter != "" else " "
                parts = [p.strip() for p in text.split(sep)]
            max_parts = max(max_parts, len(parts))
            row_splits[r] = parts

        if norm_mode == "split_vn_fullname":
            default_names = ["Họ và tên đệm", "Tên"] if (target_columns and len(target_columns) == 2) else ["Họ", "Tên đệm", "Tên"]
            col_names = list(target_columns) if target_columns else default_names
        else:
            count = len(target_columns) if target_columns else max_parts
            col_names = (
                list(target_columns)
                if target_columns
                else [f"{src_header_val}_{i + 1}" for i in range(count)]
            )

        start_col = max_col + 1
        for offset, cname in enumerate(col_names):
            hdr_cell = ws.cell(header_row, start_col + offset, value=sanitize_xml_string(cname))
            _clone_cell_style(proto_header, hdr_cell)
            created_columns.append(cname)

        for r, parts in row_splits.items():
            rows_transformed += 1
            for offset in range(len(col_names)):
                val = parts[offset] if offset < len(parts) else ""
                dst_cell = ws.cell(r, start_col + offset, value=sanitize_xml_string(val))
                _clone_cell_style(proto_data, dst_cell)
                if val.isdigit() and val.startswith("0") and len(val) > 1:
                    dst_cell.data_type = "s"
                    dst_cell.number_format = "@"

    _auto_fit_worksheet_columns(ws, header_row=header_row)
    apply_worksheet_print_setup(ws, print_title_header_row=header_row)
    wb.save(str(dst))

    return {
        "ok": True,
        "mode": norm_mode,
        "input_path": str(src),
        "output_path": str(dst),
        "sheet_name": ws.title,
        "rows_transformed": rows_transformed,
        "created_columns": created_columns,
    }


# ---------------------------------------------------------------------------
# 3. Pivot Summary + Embedded Native Chart (Live Cross-Sheet Formulas)
# ---------------------------------------------------------------------------


def build_pivot_summary_and_chart(
    input_path: str,
    group_by_col: str,
    value_cols: Optional[Sequence[str]] = None,
    agg_func: str = "SUM",
    chart_type: str = "bar",
    output_path: Optional[str] = None,
    source_sheet: Optional[str] = None,
    summary_sheet_name: str = "Pivot_Summary",
) -> Dict[str, Any]:
    """
    Generate an in-workbook Pivot Summary tab driven by live Excel formulas
    (`SUMIFS`, `COUNTIFS`, `AVERAGEIFS`) plus a native embedded chart (`bar`, `line`, `pie`).
    """
    src = _ensure_existing_file(input_path)
    dst = _resolve_output_path(src, output_path, "pivot")
    norm_agg = (agg_func or "SUM").strip().upper()
    if norm_agg not in {"SUM", "COUNT", "AVERAGE"}:
        raise ValueError("agg_func must be one of: SUM, COUNT, AVERAGE.")
    norm_chart = (chart_type or "bar").strip().lower()
    if norm_chart not in {"bar", "line", "pie"}:
        raise ValueError("chart_type must be one of: bar, line, pie.")

    wb = openpyxl.load_workbook(str(src))
    src_ws = _select_worksheet(wb, source_sheet)
    src_header_row = _detect_header_row(src_ws)
    headers, _ = _get_header_map(src_ws, src_header_row)
    src_max_row = src_ws.max_row or src_header_row
    group_idx = _resolve_column_index(group_by_col, src_ws, src_header_row)
    group_header = headers[group_idx - 1] if group_idx - 1 < len(headers) else group_by_col

    # Auto-pick first numeric column if value_cols omitted for SUM/AVERAGE
    val_indices: List[int] = []
    if value_cols:
        val_indices = [_resolve_column_index(vc, src_ws, src_header_row) for vc in value_cols]
    elif norm_agg != "COUNT":
        sample_row = min(src_header_row + 1, src_max_row)
        for c in range(1, (src_ws.max_column or 1) + 1):
            if c == group_idx:
                continue
            v = src_ws.cell(sample_row, c).value
            if isinstance(v, (int, float)) and not isinstance(v, bool):
                val_indices.append(c)
                break
        if not val_indices:
            norm_agg = "COUNT"

    # Extract distinct ordered groups
    seen_groups: Dict[str, Any] = {}
    for r in range(src_header_row + 1, src_max_row + 1):
        raw = src_ws.cell(r, group_idx).value
        if raw is None or str(raw).strip() == "":
            continue
        key = _normalize_key_for_join(raw)
        if key not in seen_groups:
            seen_groups[key] = raw

    if not seen_groups:
        raise ValueError(f"No non-empty group values found in column '{group_header}'.")

    # Recreate summary sheet cleanly while preserving source workbook sheets
    safe_sheet = sanitize_xml_string(summary_sheet_name)[:31] or "Pivot_Summary"
    if safe_sheet in wb.sheetnames:
        del wb[safe_sheet]
    piv_ws = wb.create_sheet(title=safe_sheet)
    piv_ws.views.sheetView[0].showGridLines = True

    # Title banner (Row 1) + metadata subtitle (Row 2)
    piv_ws["A1"] = f"TỔNG HỢP PIVOT THEO {group_header.upper()} ({norm_agg})"
    piv_ws["A1"].font = Font(name="Calibri", size=14, bold=True, color="1E3A5F")
    piv_ws["A2"] = f"Nguồn: '{src_ws.title}' | Dòng dữ liệu: {src_header_row + 1}..{src_max_row}"
    piv_ws["A2"].font = Font(name="Calibri", size=10, italic=True, color="475569")

    # Header row at Row 4
    tbl_hdr_row = 4
    metric_headers: List[str] = []
    if norm_agg == "COUNT":
        metric_headers = ["Số lượng (COUNT)"]
    else:
        for idx in val_indices:
            vh = headers[idx - 1] if idx - 1 < len(headers) else get_column_letter(idx)
            metric_headers.append(f"{norm_agg}({vh})")

    all_headers = [group_header, *metric_headers, "Tỷ trọng (%)"]
    for col_pos, title in enumerate(all_headers, start=1):
        cell = piv_ws.cell(tbl_hdr_row, col_pos, value=sanitize_xml_string(title))
        cell.fill = _HEADER_FILL
        cell.font = _HEADER_FONT
        cell.border = _HEADER_BORDER
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    escaped_src = src_ws.title.replace("'", "''")
    group_col_letter = get_column_letter(group_idx)
    crit_range = f"'{escaped_src}'!${group_col_letter}${src_header_row + 1}:${group_col_letter}${src_max_row}"

    data_start_row = tbl_hdr_row + 1
    total_groups = len(seen_groups)
    total_row = data_start_row + total_groups

    for offset, (norm_key, orig_val) in enumerate(seen_groups.items()):
        r = data_start_row + offset
        key_cell = piv_ws.cell(r, 1, value=orig_val)
        key_cell.font = Font(name="Calibri", size=11, color="0F172A")
        key_cell.border = _CELL_BORDER
        if isinstance(orig_val, str) and orig_val.isdigit() and orig_val.startswith("0"):
            key_cell.number_format = "@"

        crit_ref = f"$A{r}"
        if norm_agg == "COUNT":
            formula = f"=COUNTIFS({crit_range},{crit_ref})"
            c_metric = piv_ws.cell(r, 2, value=formula)
            c_metric.number_format = "#,##0"
            c_metric.border = _CELL_BORDER
            c_metric.alignment = Alignment(horizontal="right", vertical="center")
        else:
            for m_idx, src_v_idx in enumerate(val_indices, start=2):
                v_letter = get_column_letter(src_v_idx)
                sum_range = f"'{escaped_src}'!${v_letter}${src_header_row + 1}:${v_letter}${src_max_row}"
                fn = "SUMIFS" if norm_agg == "SUM" else "AVERAGEIFS"
                formula = f"={fn}({sum_range},{crit_range},{crit_ref})"
                c_metric = piv_ws.cell(r, m_idx, value=formula)
                c_metric.number_format = "#,##0.##"
                c_metric.border = _CELL_BORDER
                c_metric.alignment = Alignment(horizontal="right", vertical="center")

        # Share percentage column based on first metric column (Col B)
        share_col_idx = len(all_headers)
        share_cell = piv_ws.cell(r, share_col_idx, value=f"=IFERROR(B{r}/$B${total_row},0)")
        share_cell.number_format = "0.0%"
        share_cell.border = _CELL_BORDER
        share_cell.alignment = Alignment(horizontal="right", vertical="center")

        if offset % 2 == 1:
            for c_idx in range(1, len(all_headers) + 1):
                piv_ws.cell(r, c_idx).fill = _ZEBRA_FILL

    # Total row (E3 + E14)
    lbl_total = piv_ws.cell(total_row, 1, value="Tổng cộng")
    lbl_total.fill = _SUBTOTAL_FILL
    lbl_total.font = _SUBTOTAL_FONT
    lbl_total.border = _CELL_BORDER
    for c_idx in range(2, len(all_headers)):
        col_l = get_column_letter(c_idx)
        agg_formula = (
            f"=AVERAGE({col_l}{data_start_row}:{col_l}{total_row - 1})"
            if norm_agg == "AVERAGE"
            else f"=SUM({col_l}{data_start_row}:{col_l}{total_row - 1})"
        )
        t_cell = piv_ws.cell(total_row, c_idx, value=agg_formula)
        t_cell.fill = _SUBTOTAL_FILL
        t_cell.font = _SUBTOTAL_FONT
        t_cell.border = _CELL_BORDER
        t_cell.number_format = "#,##0" if norm_agg == "COUNT" else "#,##0.##"
        t_cell.alignment = Alignment(horizontal="right", vertical="center")

    share_tot = piv_ws.cell(total_row, len(all_headers), value=f"=SUM({get_column_letter(len(all_headers))}{data_start_row}:{get_column_letter(len(all_headers))}{total_row - 1})")
    share_tot.fill = _SUBTOTAL_FILL
    share_tot.font = _SUBTOTAL_FONT
    share_tot.border = _CELL_BORDER
    share_tot.number_format = "0.0%"
    share_tot.alignment = Alignment(horizontal="right", vertical="center")

    # Embedded Native Chart referencing live Pivot range (E12)
    chart: Union[BarChart, LineChart, PieChart]
    if norm_chart == "pie":
        chart = PieChart()
        data_ref = Reference(piv_ws, min_col=2, min_row=tbl_hdr_row, max_row=total_row - 1)
    elif norm_chart == "line":
        chart = LineChart()
        data_ref = Reference(
            piv_ws, min_col=2, max_col=1 + len(metric_headers), min_row=tbl_hdr_row, max_row=total_row - 1
        )
    else:
        chart = BarChart()
        chart.type = "col"
        chart.grouping = "clustered"
        data_ref = Reference(
            piv_ws, min_col=2, max_col=1 + len(metric_headers), min_row=tbl_hdr_row, max_row=total_row - 1
        )

    cats_ref = Reference(piv_ws, min_col=1, min_row=data_start_row, max_row=total_row - 1)
    chart.title = f"{group_header} ({norm_agg})"
    chart.style = 10
    chart.width = 18
    chart.height = 10
    chart.add_data(data_ref, titles_from_data=True)
    chart.set_categories(cats_ref)

    anchor_col = get_column_letter(len(all_headers) + 2)
    piv_ws.add_chart(chart, f"{anchor_col}{tbl_hdr_row}")

    _auto_fit_worksheet_columns(piv_ws, header_row=tbl_hdr_row)
    piv_ws.freeze_panes = f"A{data_start_row}"
    apply_worksheet_print_setup(piv_ws, orientation="landscape", print_title_header_row=tbl_hdr_row)
    wb.save(str(dst))

    return {
        "ok": True,
        "input_path": str(src),
        "output_path": str(dst),
        "source_sheet": src_ws.title,
        "summary_sheet": safe_sheet,
        "groups_count": total_groups,
        "agg_func": norm_agg,
        "chart_type": norm_chart,
        "value_columns": metric_headers,
    }


# ---------------------------------------------------------------------------
# 4. No-Code XLOOKUP Join
# ---------------------------------------------------------------------------


def xlookup_join_workbooks(
    left_path: str,
    left_key: str,
    right_key: str,
    return_cols: Sequence[str],
    right_path: Optional[str] = None,
    output_path: Optional[str] = None,
    left_sheet: Optional[str] = None,
    right_sheet: Optional[str] = None,
    not_found_value: str = "N/A",
    formula_mode: str = "auto",
) -> Dict[str, Any]:
    """
    Perform a no-code VLOOKUP/XLOOKUP style join merging `return_cols` onto the left sheet.
    - Same-workbook joins emit live `=IFERROR(INDEX(...,MATCH(...)), ...)` formulas by default.
    - Cross-workbook joins emit direct normalized values (no broken external `#REF!` links)
      unless `formula_mode="formula"`, which copies the lookup tab cleanly inside the workbook.
    """
    if not return_cols:
        raise ValueError("return_cols must specify at least one column to bring over.")

    left_src = _ensure_existing_file(left_path, "Left workbook")
    right_src = _ensure_existing_file(right_path or left_path, "Right workbook")
    dst = _resolve_output_path(left_src, output_path, "joined")

    same_workbook = left_src.resolve() == right_src.resolve()
    left_wb = openpyxl.load_workbook(str(left_src))
    left_ws = _select_worksheet(left_wb, left_sheet)

    right_wb = left_wb if same_workbook else openpyxl.load_workbook(str(right_src), data_only=True)
    right_ws = _select_worksheet(right_wb, right_sheet)

    l_hdr_row = _detect_header_row(left_ws)
    r_hdr_row = _detect_header_row(right_ws)
    r_headers, _ = _get_header_map(right_ws, r_hdr_row)

    l_key_idx = _resolve_column_index(left_key, left_ws, l_hdr_row)
    r_key_idx = _resolve_column_index(right_key, right_ws, r_hdr_row)
    r_ret_indices = [_resolve_column_index(col, right_ws, r_hdr_row) for col in return_cols]

    # Build lookup index from right sheet (first-match semantics like XLOOKUP)
    r_max_row = right_ws.max_row or r_hdr_row
    lookup_index: Dict[str, Dict[int, Any]] = {}
    for r in range(r_hdr_row + 1, r_max_row + 1):
        key_val = _normalize_key_for_join(right_ws.cell(r, r_key_idx).value)
        if not key_val or key_val in lookup_index:
            continue
        lookup_index[key_val] = {c_idx: right_ws.cell(r, c_idx).value for c_idx in r_ret_indices}

    use_live_formula = (formula_mode == "formula") or (formula_mode == "auto" and same_workbook)
    if use_live_formula and not same_workbook:
        # Copy right table into a clean internal lookup sheet so formulas have 0 external links
        ref_sheet_name = "_Lookup_Ref"
        if ref_sheet_name in left_wb.sheetnames:
            del left_wb[ref_sheet_name]
        ref_ws = left_wb.create_sheet(title=ref_sheet_name)
        for r in range(1, r_max_row + 1):
            for c in range(1, (right_ws.max_column or 1) + 1):
                ref_ws.cell(r, c, value=right_ws.cell(r, c).value)
        right_ws = ref_ws

    # Append joined columns cleanly preserving left sheet prototype header/cell styles
    l_max_col = left_ws.max_column or 1
    l_max_row = left_ws.max_row or l_hdr_row
    proto_hdr = left_ws.cell(l_hdr_row, l_key_idx)
    proto_data = left_ws.cell(min(l_hdr_row + 1, l_max_row), l_key_idx)

    added_columns: List[str] = []
    escaped_right = right_ws.title.replace("'", "''")
    r_key_letter = get_column_letter(r_key_idx)
    l_key_letter = get_column_letter(l_key_idx)
    escaped_not_found = str(not_found_value).replace('"', '""')

    for offset, r_col_idx in enumerate(r_ret_indices, start=1):
        dest_col = l_max_col + offset
        col_title = r_headers[r_col_idx - 1] if r_col_idx - 1 < len(r_headers) else get_column_letter(r_col_idx)
        added_columns.append(col_title)
        hdr_cell = left_ws.cell(l_hdr_row, dest_col, value=sanitize_xml_string(col_title))
        _clone_cell_style(proto_hdr, hdr_cell)

    matched_rows = 0
    unmatched_rows = 0
    for r in range(l_hdr_row + 1, l_max_row + 1):
        raw_left_key = left_ws.cell(r, l_key_idx).value
        norm_k = _normalize_key_for_join(raw_left_key)
        if not norm_k:
            continue
        match_record = lookup_index.get(norm_k)
        if match_record is not None:
            matched_rows += 1
        else:
            unmatched_rows += 1

        for offset, r_col_idx in enumerate(r_ret_indices, start=1):
            dest_col = l_max_col + offset
            is_id, _ = _is_id_or_phone_header(added_columns[offset - 1])
            if use_live_formula:
                ret_letter = get_column_letter(r_col_idx)
                r_key_rng = f"'{escaped_right}'!${r_key_letter}${r_hdr_row + 1}:${r_key_letter}${r_max_row}"
                r_val_rng = f"'{escaped_right}'!${ret_letter}${r_hdr_row + 1}:${ret_letter}${r_max_row}"
                formula = (
                    f'=IFERROR(INDEX({r_val_rng},MATCH(${l_key_letter}{r},{r_key_rng},0)),"{escaped_not_found}")'
                )
                dst_cell = left_ws.cell(r, dest_col, value=formula)
                _clone_cell_style(proto_data, dst_cell)
            else:
                val = match_record[r_col_idx] if match_record is not None else not_found_value
                dst_cell = left_ws.cell(r, dest_col)
                _clone_cell_style(proto_data, dst_cell)
                if is_id and val is not None and val != not_found_value:
                    dst_cell.value = sanitize_xml_string(_normalize_identifier_value(val, is_phone=False))
                    dst_cell.data_type = "s"
                    dst_cell.number_format = "@"
                else:
                    dst_cell.value = sanitize_xml_string(val) if isinstance(val, str) else val

    _auto_fit_worksheet_columns(left_ws, header_row=l_hdr_row)
    apply_worksheet_print_setup(left_ws, print_title_header_row=l_hdr_row)
    left_wb.save(str(dst))

    return {
        "ok": True,
        "output_path": str(dst),
        "left_sheet": left_ws.title,
        "right_sheet": right_ws.title,
        "matched_rows": matched_rows,
        "unmatched_rows": unmatched_rows,
        "added_columns": added_columns,
        "live_formulas": use_live_formula,
    }


# ---------------------------------------------------------------------------
# 5. Split Master Workbook by Column Value -> Separate XLSX + ZIP Bundle
# ---------------------------------------------------------------------------


def split_workbook_by_column(
    input_path: str,
    split_col: str,
    output_path: Optional[str] = None,
    sheet_name: Optional[str] = None,
    as_zip: bool = True,
) -> Dict[str, Any]:
    """
    Split a master worksheet into separate Excel files grouped by unique values
    of `split_col`, preserving headers, cell styles, number formats, and print setup.
    Returns paths / ZIP bundle metadata.
    """
    src = _ensure_existing_file(input_path)
    src_wb = openpyxl.load_workbook(str(src))
    src_ws = _select_worksheet(src_wb, sheet_name)

    header_row = _detect_header_row(src_ws)
    split_idx = _resolve_column_index(split_col, src_ws, header_row)
    max_col = src_ws.max_column or 1
    max_row = src_ws.max_row or header_row

    groups: Dict[str, List[int]] = {}
    for r in range(header_row + 1, max_row + 1):
        row_vals = [src_ws.cell(r, c).value for c in range(1, max_col + 1)]
        if all(v is None or str(v).strip() == "" for v in row_vals):
            continue
        raw_key = src_ws.cell(r, split_idx).value
        group_name = str(raw_key).strip() if raw_key is not None and str(raw_key).strip() else "Unassigned"
        groups.setdefault(group_name, []).append(r)

    if not groups:
        raise ValueError(f"No data rows found to split in sheet '{src_ws.title}'.")

    if as_zip:
        out_zip = (
            Path(output_path)
            if output_path
            else src.with_name(f"{src.stem}_split.zip")
        )
        out_zip.parent.mkdir(parents=True, exist_ok=True)
        work_dir = Path(tempfile.mkdtemp(prefix="antigravity_split_"))
    else:
        work_dir = Path(output_path) if output_path else src.with_name(f"{src.stem}_split")
        work_dir.mkdir(parents=True, exist_ok=True)
        out_zip = work_dir

    generated_files: List[Dict[str, Any]] = []
    used_filenames: set[str] = set()

    for group_key, row_indices in groups.items():
        slug = _safe_filename_slug(group_key, fallback="Group")
        candidate = f"{slug}.xlsx"
        suffix_num = 2
        while candidate.lower() in used_filenames:
            candidate = f"{slug}_{suffix_num}.xlsx"
            suffix_num += 1
        used_filenames.add(candidate.lower())

        child_wb = openpyxl.Workbook()
        child_ws = child_wb.active
        child_ws.title = src_ws.title

        # Copy column widths
        for c in range(1, max_col + 1):
            col_l = get_column_letter(c)
            if col_l in src_ws.column_dimensions and src_ws.column_dimensions[col_l].width:
                child_ws.column_dimensions[col_l].width = src_ws.column_dimensions[col_l].width

        # Copy header rows (1..header_row) + matched data rows
        rows_to_copy = list(range(1, header_row + 1)) + row_indices
        for target_r, orig_r in enumerate(rows_to_copy, start=1):
            if orig_r in src_ws.row_dimensions and src_ws.row_dimensions[orig_r].height:
                child_ws.row_dimensions[target_r].height = src_ws.row_dimensions[orig_r].height
            for c in range(1, max_col + 1):
                src_c = src_ws.cell(orig_r, c)
                dst_c = child_ws.cell(target_r, c, value=src_c.value)
                _clone_cell_style(src_c, dst_c)
                if src_c.data_type == "s":
                    dst_c.data_type = "s"

        child_ws.freeze_panes = child_ws.cell(header_row + 1, 1).coordinate
        apply_worksheet_print_setup(child_ws, print_title_header_row=header_row)
        child_file = work_dir / candidate
        child_wb.save(str(child_file))
        generated_files.append(
            {
                "group": group_key,
                "filename": candidate,
                "path": str(child_file),
                "rows": len(row_indices),
            }
        )

    if as_zip:
        with zipfile.ZipFile(str(out_zip), "w", compression=zipfile.ZIP_DEFLATED) as zf:
            for item in generated_files:
                zf.write(item["path"], arcname=item["filename"])

    return {
        "ok": True,
        "input_path": str(src),
        "output_path": str(out_zip),
        "as_zip": as_zip,
        "total_files": len(generated_files),
        "total_rows": sum(item["rows"] for item in generated_files),
        "groups": generated_files,
    }


# ---------------------------------------------------------------------------
# 6. Visual Excel Diff (Cell Highlighting + Comments + _Diff_Summary Sheet)
# ---------------------------------------------------------------------------


def _values_equal(v1: Any, v2: Any) -> bool:
    """Compare cell values cleanly accounting for int/float precision and None vs ''."""
    n1 = "" if v1 is None else v1
    n2 = "" if v2 is None else v2
    if n1 == n2:
        return True
    if isinstance(n1, (int, float)) and isinstance(n2, (int, float)):
        return abs(float(n1) - float(n2)) < 1e-9
    return str(n1).strip() == str(n2).strip()


def diff_workbooks_visual(
    v1_path: str,
    v2_path: str,
    output_path: Optional[str] = None,
    sheets: Optional[Sequence[str]] = None,
) -> Dict[str, Any]:
    """
    Compare v1 vs v2 Excel workbooks cell-by-cell and write an annotated copy of v2
    with Amber (modified), Emerald (added), and Rose (removed) highlights,
    Comment tooltips showing old values, and a clickable `_Diff_Summary` tab.
    """
    src_v1 = _ensure_existing_file(v1_path, "Baseline workbook (v1)")
    src_v2 = _ensure_existing_file(v2_path, "Target workbook (v2)")
    dst = _resolve_output_path(src_v2, output_path, "diff")

    wb1 = openpyxl.load_workbook(str(src_v1), data_only=False)
    wb2 = openpyxl.load_workbook(str(src_v2), data_only=False)

    target_sheets = list(sheets) if sheets else list(wb2.sheetnames)
    changes: List[Dict[str, Any]] = []
    sheet_summaries: List[Dict[str, int]] = []
    modified_cells = 0
    added_cells = 0
    removed_cells = 0

    for s_name in target_sheets:
        if s_name == "_Diff_Summary":
            continue
        if s_name not in wb2.sheetnames:
            continue
        ws2 = wb2[s_name]
        ws1 = wb1[s_name] if s_name in wb1.sheetnames else None

        s_mod = 0
        s_add = 0
        s_rem = 0
        max_r = max(ws2.max_row or 1, (ws1.max_row or 1) if ws1 else 1)
        max_c = max(ws2.max_column or 1, (ws1.max_column or 1) if ws1 else 1)

        for r in range(1, max_r + 1):
            for c in range(1, max_c + 1):
                val1 = ws1.cell(r, c).value if ws1 else None
                val2 = ws2.cell(r, c).value
                if _values_equal(val1, val2):
                    continue

                coord = f"{get_column_letter(c)}{r}"
                cell2 = ws2.cell(r, c)
                if val1 is None or str(val1).strip() == "":
                    kind = "ADDED"
                    cell2.fill = _DIFF_ADDED_FILL
                    cell2.font = _DIFF_ADDED_FONT
                    cell2.comment = Comment("Added in v2 (was empty in v1)", "Antigravity Diff")
                    added_cells += 1
                    s_add += 1
                elif val2 is None or str(val2).strip() == "":
                    kind = "REMOVED"
                    cell2.fill = _DIFF_REMOVED_FILL
                    cell2.font = _DIFF_REMOVED_FONT
                    cell2.comment = Comment(f"Removed in v2 (v1 was: {val1})", "Antigravity Diff")
                    removed_cells += 1
                    s_rem += 1
                else:
                    kind = "MODIFIED"
                    cell2.fill = _DIFF_MODIFIED_FILL
                    cell2.font = _DIFF_MODIFIED_FONT
                    cell2.comment = Comment(f"v1: {val1}\nv2: {val2}", "Antigravity Diff")
                    modified_cells += 1
                    s_mod += 1

                changes.append(
                    {
                        "sheet": s_name,
                        "cell": coord,
                        "type": kind,
                        "v1": "" if val1 is None else str(val1),
                        "v2": "" if val2 is None else str(val2),
                    }
                )

        sheet_summaries.append(
            {"sheet": s_name, "modified": s_mod, "added": s_add, "removed": s_rem}
        )

    # Check sheets completely deleted from v1
    for s_name in wb1.sheetnames:
        if s_name not in wb2.sheetnames and (not sheets or s_name in sheets):
            changes.append(
                {
                    "sheet": s_name,
                    "cell": "ALL",
                    "type": "REMOVED_SHEET",
                    "v1": f"Sheet '{s_name}' existed",
                    "v2": "(Deleted in v2)",
                }
            )

    # Build _Diff_Summary sheet at index 0
    if "_Diff_Summary" in wb2.sheetnames:
        del wb2["_Diff_Summary"]
    sum_ws = wb2.create_sheet(title="_Diff_Summary", index=0)
    sum_ws["A1"] = "BÁO CÁO ĐỐI CHIẾU EXCEL DIFF (v1 -> v2)"
    sum_ws["A1"].font = Font(name="Calibri", size=14, bold=True, color="1E3A5F")
    sum_ws["A2"] = (
        f"Modified: {modified_cells} | Added: {added_cells} | Removed: {removed_cells} | "
        f"Total Diffs: {len(changes)}"
    )
    sum_ws["A2"].font = Font(name="Calibri", size=10, bold=True, color="334155")

    hdr_row = 4
    headers = ["Sheet", "Cell", "Loại thay đổi", "Giá trị cũ (v1)", "Giá trị mới (v2)"]
    for idx, h in enumerate(headers, start=1):
        hc = sum_ws.cell(hdr_row, idx, value=h)
        hc.fill = _HEADER_FILL
        hc.font = _HEADER_FONT
        hc.border = _HEADER_BORDER
        hc.alignment = Alignment(horizontal="center", vertical="center")

    for offset, item in enumerate(changes[:2000], start=1):
        r = hdr_row + offset
        sum_ws.cell(r, 1, value=item["sheet"]).border = _CELL_BORDER
        cell_ref = sum_ws.cell(r, 2)
        if item["cell"] != "ALL" and item["sheet"] in wb2.sheetnames:
            esc = item["sheet"].replace("'", "''")
            cell_ref.value = f'=HYPERLINK("#\'{esc}\'!{item["cell"]}", "{item["cell"]}")'
            cell_ref.font = Font(name="Calibri", size=11, underline="single", color="1D4ED8")
        else:
            cell_ref.value = item["cell"]
        cell_ref.border = _CELL_BORDER

        kind_cell = sum_ws.cell(r, 3, value=item["type"])
        kind_cell.border = _CELL_BORDER
        if item["type"] == "MODIFIED":
            kind_cell.fill = _DIFF_MODIFIED_FILL
            kind_cell.font = _DIFF_MODIFIED_FONT
        elif item["type"] == "ADDED":
            kind_cell.fill = _DIFF_ADDED_FILL
            kind_cell.font = _DIFF_ADDED_FONT
        else:
            kind_cell.fill = _DIFF_REMOVED_FILL
            kind_cell.font = _DIFF_REMOVED_FONT

        v1_c = sum_ws.cell(r, 4, value=sanitize_xml_string(item["v1"]))
        v1_c.border = _CELL_BORDER
        v1_c.number_format = "@"
        v2_c = sum_ws.cell(r, 5, value=sanitize_xml_string(item["v2"]))
        v2_c.border = _CELL_BORDER
        v2_c.number_format = "@"

    _auto_fit_worksheet_columns(sum_ws, header_row=hdr_row)
    sum_ws.freeze_panes = "A5"
    apply_worksheet_print_setup(sum_ws, orientation="landscape", print_title_header_row=hdr_row)
    wb2.save(str(dst))

    return {
        "ok": True,
        "v1_path": str(src_v1),
        "v2_path": str(src_v2),
        "output_path": str(dst),
        "total_changes": len(changes),
        "modified_cells": modified_cells,
        "added_cells": added_cells,
        "removed_cells": removed_cells,
        "sheet_summaries": sheet_summaries,
        "sample_changes": changes[:50],
    }
