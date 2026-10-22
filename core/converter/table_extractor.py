"""
core.converter.table_extractor — Semantic Table Extraction Engine (Universal 5-Way Matrix).

Extracts structured tabular data from PDF, DOCX, and Markdown into enterprise-grade
Excel (.xlsx) workbooks with high-fidelity formatting:
1. DOCX: Scans OpenXML <w:tbl>, filters administrative layout tables (Decree 30),
   resolves <w:gridSpan> (colspan) and <w:vMerge> (rowspan) safely on XML level without
   triggering openpyxl MergedCell read-only errors (ERR_XLSX_004).
2. PDF: Utilizes PyMuPDF's C++ table finder (page.find_tables()) for layout extraction.
3. Markdown: Parses GFM table blocks (| Col 1 | Col 2 |) and maps nearest # Headings.
4. Styling: Navy headers (#0F294A), zebra striping (#F8FAFC), string formatting (@)
   for numeric IDs/codes (ERR_XLSX_007), and landscape A4 print setup (ERR_XLSX_008).
"""

from __future__ import annotations

import os
import re
from typing import Any, Dict, List, Optional, Set, Tuple

try:
    import openpyxl
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter
except ImportError:
    openpyxl = None  # type: ignore

try:
    import docx
    from docx.oxml.ns import qn
    from docx.table import Table
except ImportError:
    docx = None  # type: ignore

try:
    import pymupdf as fitz
except ImportError:
    try:
        import fitz  # type: ignore
    except ImportError:
        fitz = None  # type: ignore

try:
    from core.xlsx.writer import apply_worksheet_print_setup
except ImportError:
    apply_worksheet_print_setup = None  # type: ignore


# ---------------------------------------------------------------------------
# Visual Tokens & Styles
# ---------------------------------------------------------------------------

HEADER_FILL = PatternFill(start_color="0F294A", end_color="0F294A", fill_type="solid") if openpyxl else None
HEADER_FONT = Font(name="Calibri", size=11, bold=True, color="FFFFFF") if openpyxl else None
DATA_FONT = Font(name="Calibri", size=10) if openpyxl else None
ALT_FILL = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid") if openpyxl else None
WHITE_FILL = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid") if openpyxl else None

THIN_BORDER = Border(
    left=Side(style="thin", color="CBD5E1"),
    right=Side(style="thin", color="CBD5E1"),
    top=Side(style="thin", color="CBD5E1"),
    bottom=Side(style="thin", color="CBD5E1"),
) if openpyxl else None


# ---------------------------------------------------------------------------
# Helper: Sheet Title Sanitization (Edge Case 3)
# ---------------------------------------------------------------------------

def _sanitize_sheet_title(raw_title: str, existing_names: Set[str], fallback_idx: int) -> str:
    """
    Sanitizes sheet titles for Excel:
    - Max 31 characters
    - Removes forbidden characters: \\ / ? * : [ ]
    - Strips leading/trailing spaces and quotes
    - Ensures uniqueness across existing_names (case-insensitive check)
    """
    clean = re.sub(r'[\\/?*:[\]]', '_', raw_title or "").strip(" '\"\t\r\n")
    clean = re.sub(r'\s+', '_', clean)
    clean = clean[:28].strip("_")
    if not clean:
        clean = f"Bang_{fallback_idx}"

    candidate = clean
    suffix = 1
    lower_existing = {name.lower() for name in existing_names}
    while candidate.lower() in lower_existing:
        suffix += 1
        base = clean[:25].rstrip("_")
        candidate = f"{base}_{suffix}"
        if len(candidate) > 31:
            candidate = f"B_{fallback_idx}_{suffix}"[:31]

    existing_names.add(candidate)
    return candidate


# ---------------------------------------------------------------------------
# Helper: Decree 30 Layout Table Detection (Edge Case 1)
# ---------------------------------------------------------------------------

def _is_layout_header_or_signature_table(table: Any) -> bool:
    """
    Detects if a table is purely an administrative layout structure
    (e.g. Decree 30 National Header or Signature/Recipient block),
    which should NOT be extracted to Excel data sheets.
    """
    rows_count = len(table.rows)
    cols_count = len(table.columns) if table.rows else 0
    if rows_count > 4 or cols_count > 3:
        return False

    full_text = " ".join(
        cell.text.strip() for row in table.rows for cell in row.cells
    ).lower()

    administrative_markers = [
        "cộng hòa xã hội chủ nghĩa việt nam",
        "độc lập - tự do - hạnh phúc",
        "độc lập – tự do – hạnh phúc",
        "nơi nhận:",
        "thủ trưởng đơn vị",
        "thẩm quyền ký",
        "người ký:",
        "quyết định số:",
        "bộ thông tin và truyền thông",
    ]

    for marker in administrative_markers:
        if marker in full_text:
            return True

    return False


# ---------------------------------------------------------------------------
# Helper: Cell Value & Format Setup (ERR_XLSX_007)
# ---------------------------------------------------------------------------

def _apply_cell_value_and_format(
    cell: Any,
    val_str: str,
    is_header: bool = False,
    row_idx: int = 1,
) -> None:
    """
    Sets cell value, styling, and number format adhering to:
    - ERR_XLSX_007: String '@' format for codes, CCCD, tax IDs, phone numbers.
    - Right-align for pure numbers, Left-align for text, Center for headers.
    """
    val_clean = val_str.strip()

    is_leading_zero_code = bool(re.match(r"^0\d+$", val_clean)) and len(val_clean) > 1
    is_card_or_cccd = bool(re.match(r"^\d{9,19}$", val_clean))

    if is_leading_zero_code or is_card_or_cccd:
        cell.number_format = "@"
        cell.value = val_clean
    elif not is_header and re.match(r"^-?\d+(?:\.\d+)?$", val_clean):
        try:
            cell.value = float(val_clean) if "." in val_clean else int(val_clean)
        except ValueError:
            cell.value = val_clean
    else:
        cell.value = val_clean

    if is_header:
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    else:
        cell.fill = ALT_FILL if row_idx % 2 == 0 else WHITE_FILL
        cell.font = DATA_FONT
        align_h = "right" if isinstance(cell.value, (int, float)) else "left"
        cell.alignment = Alignment(horizontal=align_h, vertical="center", wrap_text=True)

    cell.border = THIN_BORDER


def _format_merged_cell_range(
    ws: Any,
    start_row: int,
    start_col: int,
    end_row: int,
    end_col: int,
    is_header: bool = False,
) -> None:
    """
    Synchronizes borders and fills across all cells in a merged range
    without setting .value (strictly enforcing ERR_XLSX_004).
    """
    for r in range(start_row, end_row + 1):
        for c in range(start_col, end_col + 1):
            cell = ws.cell(row=r, column=c)
            cell.border = THIN_BORDER
            if is_header:
                cell.fill = HEADER_FILL
                cell.font = HEADER_FONT
            else:
                cell.fill = ALT_FILL if start_row % 2 == 0 else WHITE_FILL
                cell.font = DATA_FONT


# ---------------------------------------------------------------------------
# DOCX Table Extractor (Edge Case 2: XML <w:tc> iteration)
# ---------------------------------------------------------------------------

def _map_docx_tables_to_captions(doc: Any) -> List[Tuple[Any, str]]:
    """
    Traverses the document body in XML order to associate each <w:tbl>
    with the nearest preceding Heading or Table Caption.
    """
    table_bindings: List[Tuple[Any, str]] = []
    current_caption = ""
    table_index = 0

    re_caption = re.compile(
        r"^(?:Bảng|Table|Tab\.)\s*([0-9IVXLCDM]+[\.\-]?[0-9]*)\s*[:\.\-]?\s*(.*)",
        re.IGNORECASE
    )

    tbl_elements = {id(t._tbl): t for t in doc.tables}

    for child in doc.element.body:
        tag = child.tag
        if tag == qn("w:p"):
            p_text = "".join(child.itertext()).strip()
            if not p_text:
                continue

            # Check style
            pPr = child.find(qn("w:pPr"))
            pStyle = pPr.find(qn("w:pStyle")) if pPr is not None else None
            style_val = pStyle.get(qn("w:val"), "") if pStyle is not None else ""

            is_heading = style_val.lower().startswith("heading") or style_val.lower().startswith("tiêu đề")
            is_caption = bool(re_caption.match(p_text))

            if is_caption:
                m = re_caption.match(p_text)
                if m:
                    num_part = m.group(1).strip()
                    desc_part = m.group(2).strip()
                    current_caption = f"Bang_{num_part}_{desc_part}" if desc_part else f"Bang_{num_part}"
                else:
                    current_caption = p_text[:25]
            elif is_heading:
                clean_h = re.sub(r'^[0-9\.\s]+', '', p_text).strip()
                current_caption = clean_h[:25] if clean_h else p_text[:25]

        elif tag == qn("w:tbl"):
            table_index += 1
            tbl_obj = tbl_elements.get(id(child))
            if tbl_obj:
                caption = current_caption or f"Bang_{table_index}"
                table_bindings.append((tbl_obj, caption))
                current_caption = ""

    return table_bindings


def extract_docx_tables_to_xlsx(
    docx_path: str,
    output_path: str,
) -> Tuple[bool, str, int]:
    """
    Extracts all structured tables from Word (.docx) into an Excel workbook.
    Safely iterates XML <w:tc> to handle <w:gridSpan> and <w:vMerge> (ERR_XLSX_004).
    """
    if openpyxl is None:
        return False, "Thư viện 'openpyxl' chưa được cài đặt.", 0
    if docx is None:
        return False, "Thư viện 'python-docx' chưa được cài đặt.", 0

    src = os.path.abspath(docx_path)
    if not os.path.isfile(src):
        return False, f"Tệp Word nguồn không tồn tại: {src}", 0

    doc = docx.Document(src)
    if not doc.tables:
        return False, "Tài liệu không chứa cấu trúc bảng biểu để trích xuất sang Excel.", 0

    wb = openpyxl.Workbook()
    wb.remove(wb.active)  # Remove default blank sheet

    existing_sheet_names: Set[str] = set()
    total_extracted_tables = 0

    table_bindings = _map_docx_tables_to_captions(doc)

    for t_idx, (table, raw_caption) in enumerate(table_bindings, start=1):
        # Edge Case 1: Administrative layout table filtering
        if _is_layout_header_or_signature_table(table):
            continue

        sheet_title = _sanitize_sheet_title(raw_caption, existing_sheet_names, t_idx)
        ws = wb.create_sheet(title=sheet_title)

        col_widths: Dict[int, int] = {}
        active_vmerges: Dict[int, Dict[str, Any]] = {}

        for r_idx, row in enumerate(table.rows, start=1):
            is_header_row = (r_idx == 1)
            trPr = row._tr.find(qn("w:trPr"))
            if trPr is not None and trPr.find(qn("w:tblHeader")) is not None:
                is_header_row = True

            curr_col = 1
            tc_elements = row._tr.findall(qn("w:tc"))
            for tc in tc_elements:
                while curr_col in active_vmerges and active_vmerges[curr_col]["end_row"] >= r_idx:
                    curr_col = active_vmerges[curr_col]["end_col"] + 1

                tcPr = tc.find(qn("w:tcPr"))

                grid_span = 1
                if tcPr is not None:
                    gs_elem = tcPr.find(qn("w:gridSpan"))
                    if gs_elem is not None:
                        try:
                            grid_span = max(1, int(gs_elem.get(qn("w:val"), 1)))
                        except (ValueError, TypeError):
                            grid_span = 1

                v_merge_val: Optional[str] = None
                if tcPr is not None:
                    vm_elem = tcPr.find(qn("w:vMerge"))
                    if vm_elem is not None:
                        v_merge_val = vm_elem.get(qn("w:val"), "continue")

                p_texts = []
                for p in tc.findall(qn("w:p")):
                    t_str = "".join(t.text for t in p.iter(qn("w:t")) if t.text)
                    if t_str:
                        p_texts.append(t_str)
                cell_text = "\n".join(p_texts).strip()

                end_col = curr_col + grid_span - 1

                if v_merge_val == "continue":
                    if curr_col in active_vmerges:
                        active_vmerges[curr_col]["end_row"] = r_idx
                    curr_col = end_col + 1
                    continue

                if curr_col in active_vmerges:
                    vm_info = active_vmerges.pop(curr_col)
                    if vm_info["end_row"] > vm_info["start_row"] or vm_info["end_col"] > curr_col:
                        ws.merge_cells(
                            start_row=vm_info["start_row"],
                            start_column=curr_col,
                            end_row=vm_info["end_row"],
                            end_column=vm_info["end_col"],
                        )
                        _format_merged_cell_range(
                            ws,
                            vm_info["start_row"],
                            curr_col,
                            vm_info["end_row"],
                            vm_info["end_col"],
                            is_header=vm_info["is_header"],
                        )

                top_left_cell = ws.cell(row=r_idx, column=curr_col)
                _apply_cell_value_and_format(top_left_cell, cell_text, is_header=is_header_row, row_idx=r_idx)

                if v_merge_val == "restart":
                    active_vmerges[curr_col] = {
                        "start_row": r_idx,
                        "end_row": r_idx,
                        "end_col": end_col,
                        "is_header": is_header_row,
                    }
                elif grid_span > 1:
                    ws.merge_cells(
                        start_row=r_idx,
                        start_column=curr_col,
                        end_row=r_idx,
                        end_column=end_col,
                    )
                    _format_merged_cell_range(
                        ws,
                        r_idx,
                        curr_col,
                        r_idx,
                        end_col,
                        is_header=is_header_row,
                    )

                lines = cell_text.split("\n")
                max_line_len = max((len(l) for l in lines), default=len(cell_text))
                for c in range(curr_col, end_col + 1):
                    col_widths[c] = max(col_widths.get(c, 10), min(50, max_line_len + 3))

                curr_col = end_col + 1

        for c, vm_info in active_vmerges.items():
            if vm_info["end_row"] > vm_info["start_row"] or vm_info["end_col"] > c:
                ws.merge_cells(
                    start_row=vm_info["start_row"],
                    start_column=c,
                    end_row=vm_info["end_row"],
                    end_column=vm_info["end_col"],
                )
                _format_merged_cell_range(
                    ws,
                    vm_info["start_row"],
                    c,
                    vm_info["end_row"],
                    vm_info["end_col"],
                    is_header=vm_info["is_header"],
                )

        for c_idx, width in col_widths.items():
            col_letter = get_column_letter(c_idx)
            ws.column_dimensions[col_letter].width = max(12, width)

        ws.page_setup.paperSize = ws.PAPERSIZE_A4
        if apply_worksheet_print_setup:
            apply_worksheet_print_setup(ws, orientation="landscape")
        else:
            ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE

        total_extracted_tables += 1

    if total_extracted_tables == 0:
        return False, "Tài liệu không chứa cấu trúc bảng biểu để trích xuất sang Excel.", 0

    dest = os.path.abspath(output_path)
    out_dir = os.path.dirname(dest)
    if out_dir:
        os.makedirs(out_dir, exist_ok=True)

    wb.save(dest)
    return True, dest, total_extracted_tables


# ---------------------------------------------------------------------------
# PDF Table Extractor (Wrapping PyMuPDF find_tables)
# ---------------------------------------------------------------------------

def extract_pdf_tables_to_xlsx(
    pdf_path: str,
    output_path: str,
) -> Tuple[bool, str, int]:
    """
    Extracts all structured tables from PDF using PyMuPDF's C++ table finder.
    """
    if openpyxl is None:
        return False, "Thư viện 'openpyxl' chưa được cài đặt.", 0
    if fitz is None:
        return False, "Thư viện 'pymupdf' chưa được cài đặt.", 0

    src = os.path.abspath(pdf_path)
    if not os.path.isfile(src):
        return False, f"Tệp PDF nguồn không tồn tại: {src}", 0

    doc = fitz.open(src)
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    existing_sheet_names: Set[str] = set()
    total_extracted_tables = 0

    for pno in range(len(doc)):
        page = doc[pno]
        tabs = page.find_tables()
        if not tabs or len(tabs.tables) == 0:
            continue

        for t_idx, tab in enumerate(tabs.tables, start=1):
            raw_rows = tab.extract()
            if not raw_rows or len(raw_rows) == 0:
                continue

            raw_title = f"Trang_{pno + 1}_Bang_{t_idx}"
            sheet_title = _sanitize_sheet_title(raw_title, existing_sheet_names, total_extracted_tables + 1)
            ws = wb.create_sheet(title=sheet_title)

            col_widths: Dict[int, int] = {}

            for r_idx, row in enumerate(raw_rows, start=1):
                is_header = (r_idx == 1)

                for c_idx, val in enumerate(row, start=1):
                    val_str = str(val).strip() if val is not None else ""
                    cell = ws.cell(row=r_idx, column=c_idx)
                    _apply_cell_value_and_format(cell, val_str, is_header=is_header, row_idx=r_idx)
                    col_widths[c_idx] = max(col_widths.get(c_idx, 10), min(50, len(val_str) + 3))

            for c_idx, width in col_widths.items():
                col_letter = get_column_letter(c_idx)
                ws.column_dimensions[col_letter].width = max(12, width)

            ws.page_setup.paperSize = ws.PAPERSIZE_A4
            if apply_worksheet_print_setup:
                apply_worksheet_print_setup(ws, orientation="landscape")
            else:
                ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE

            total_extracted_tables += 1

    doc.close()

    if total_extracted_tables == 0:
        return False, "Tài liệu không chứa cấu trúc bảng biểu để trích xuất sang Excel.", 0

    dest = os.path.abspath(output_path)
    out_dir = os.path.dirname(dest)
    if out_dir:
        os.makedirs(out_dir, exist_ok=True)

    wb.save(dest)
    return True, dest, total_extracted_tables


# ---------------------------------------------------------------------------
# Markdown Table Extractor (GFM Parsing)
# ---------------------------------------------------------------------------

def extract_md_tables_to_xlsx(
    md_path: str,
    output_path: str,
) -> Tuple[bool, str, int]:
    """
    Extracts GFM Markdown tables (| Col 1 | Col 2 |) into Excel sheets.
    """
    if openpyxl is None:
        return False, "Thư viện 'openpyxl' chưa được cài đặt.", 0

    src = os.path.abspath(md_path)
    if not os.path.isfile(src):
        return False, f"Tệp Markdown nguồn không tồn tại: {src}", 0

    with open(src, "r", encoding="utf-8") as f:
        lines = f.readlines()

    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    existing_sheet_names: Set[str] = set()
    total_extracted_tables = 0

    current_heading = ""
    current_table_rows: List[List[str]] = []

    def _flush_current_table() -> None:
        nonlocal total_extracted_tables, current_table_rows, current_heading
        if not current_table_rows or len(current_table_rows) < 1:
            current_table_rows = []
            return

        total_extracted_tables += 1
        raw_title = current_heading or f"Bang_{total_extracted_tables}"
        sheet_title = _sanitize_sheet_title(raw_title, existing_sheet_names, total_extracted_tables)
        ws = wb.create_sheet(title=sheet_title)

        col_widths: Dict[int, int] = {}
        for r_idx, row in enumerate(current_table_rows, start=1):
            is_header = (r_idx == 1)
            for c_idx, val_str in enumerate(row, start=1):
                cell = ws.cell(row=r_idx, column=c_idx)
                _apply_cell_value_and_format(cell, val_str, is_header=is_header, row_idx=r_idx)
                col_widths[c_idx] = max(col_widths.get(c_idx, 10), min(50, len(val_str) + 3))

        for c_idx, width in col_widths.items():
            col_letter = get_column_letter(c_idx)
            ws.column_dimensions[col_letter].width = max(12, width)

        ws.page_setup.paperSize = ws.PAPERSIZE_A4
        if apply_worksheet_print_setup:
            apply_worksheet_print_setup(ws, orientation="landscape")
        else:
            ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE

        current_table_rows = []

    for line in lines:
        stripped = line.strip()

        if stripped.startswith("#"):
            if current_table_rows:
                _flush_current_table()
            current_heading = stripped.lstrip("#").strip()
            continue

        if stripped.startswith("|") and stripped.endswith("|"):
            inner = stripped[1:-1]
            cols = [c.strip() for c in inner.split("|")]
            if all(re.match(r"^:?-+:?$", c) for c in cols if c):
                continue

            current_table_rows.append(cols)
        else:
            if current_table_rows:
                _flush_current_table()

    if current_table_rows:
        _flush_current_table()

    if total_extracted_tables == 0:
        return False, "Tài liệu không chứa cấu trúc bảng biểu để trích xuất sang Excel.", 0

    dest = os.path.abspath(output_path)
    out_dir = os.path.dirname(dest)
    if out_dir:
        os.makedirs(out_dir, exist_ok=True)

    wb.save(dest)
    return True, dest, total_extracted_tables


# ---------------------------------------------------------------------------
# Unified Dispatcher: extract_tables_to_xlsx
# ---------------------------------------------------------------------------

def extract_tables_to_xlsx(
    source_path: str,
    output_path: str,
) -> Tuple[bool, str, int]:
    """
    Universal semantic table extractor dispatcher:
    Accepts .docx, .pdf, or .md/.markdown and routes to the appropriate parser.
    Returns:
        (success, result_path_or_error_msg, tables_extracted_count)
    """
    src = os.path.abspath(source_path)
    if not os.path.isfile(src):
        return False, f"Tệp nguồn không tồn tại: {src}", 0

    ext = os.path.splitext(src)[1].lower()
    if ext == ".docx":
        return extract_docx_tables_to_xlsx(src, output_path)
    elif ext == ".pdf":
        return extract_pdf_tables_to_xlsx(src, output_path)
    elif ext in (".md", ".markdown"):
        return extract_md_tables_to_xlsx(src, output_path)
    else:
        return False, f"Định dạng nguồn '{ext}' không được hỗ trợ để trích xuất bảng sang Excel.", 0
