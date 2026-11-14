"""
core.xlsx.tracker — Office Work & Daily Task Tracking Engine.

Refactored from tools/project_tracker.py into Layer 3 (core/xlsx/) so both the
Spreadsheet Studio Web UI, REST API, MCP Server, and CLI share one implementation
while remaining 100% schema-compatible with existing project_tracking.xlsx and
task_tracking.xlsx workbooks.
"""

from __future__ import annotations

import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

from .writer import apply_worksheet_print_setup, sanitize_xml_string

BASE_DIR = Path(__file__).resolve().parent.parent.parent
DEFAULT_PROJECT_FILE = BASE_DIR / "project_tracking.xlsx"
DEFAULT_TASK_FILE = BASE_DIR / "task_tracking.xlsx"

# Palette & Typography Tokens (Enterprise Navy/Slate matching original tracker workbooks)
_COLOR_NAVY = "1B365D"
_COLOR_SUBHEADER = "D9E1F2"
_COLOR_ZEBRA = "F9FAFC"
_COLOR_BORDER = "D1D5DB"

_FONT_TITLE = Font(name="Calibri", size=14, bold=True, color="FFFFFF")
_FONT_HDR = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
_FONT_HDR_DARK = Font(name="Calibri", size=11, bold=True, color="1B365D")
_FONT_DATA = Font(name="Calibri", size=11, bold=False, color="000000")
_FONT_DATA_BOLD = Font(name="Calibri", size=11, bold=True, color="000000")
_FONT_CODE = Font(name="Consolas", size=10, bold=True, color="0B5394")

_FILL_NAVY = PatternFill(start_color=_COLOR_NAVY, end_color=_COLOR_NAVY, fill_type="solid")
_FILL_SUBHEADER = PatternFill(
    start_color=_COLOR_SUBHEADER, end_color=_COLOR_SUBHEADER, fill_type="solid"
)
_FILL_ZEBRA = PatternFill(start_color=_COLOR_ZEBRA, end_color=_COLOR_ZEBRA, fill_type="solid")
_FILL_WHITE = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")

_THIN = Side(border_style="thin", color=_COLOR_BORDER)
_MED_NAVY = Side(border_style="medium", color=_COLOR_NAVY)
_DOUBLE_NAVY = Side(border_style="double", color=_COLOR_NAVY)

_BORDER_THIN = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)
_BORDER_HEADER = Border(left=_THIN, right=_THIN, top=_MED_NAVY, bottom=_MED_NAVY)
_BORDER_TOTAL = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_DOUBLE_NAVY)

_ALIGN_CENTER = Alignment(horizontal="center", vertical="center")
_ALIGN_LEFT = Alignment(horizontal="left", vertical="center")
_ALIGN_WRAP_CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
_ALIGN_WRAP_LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)

_LOG_HEADERS: List[Tuple[str, float]] = [
    ("STT", 6),
    ("Ngày cập nhật", 15),
    ("Module", 22),
    ("Mã tính năng / Màn hình", 18),
    ("Tên tính năng / Màn hình", 28),
    ("Phân loại", 14),
    ("Nghiệp vụ & Business Rules", 48),
    ("Thay đổi Giao diện / Mockup", 44),
    ("Trạng thái FSD", 18),
    ("Ghi chú / Lý do thay đổi", 32),
]


def _resolve_target_sheet(wb: openpyxl.Workbook) -> Worksheet:
    """Return 'FSD_Mockup_Log' if present, else fallback to first worksheet."""
    if "FSD_Mockup_Log" in wb.sheetnames:
        return wb["FSD_Mockup_Log"]
    if wb.active is None:
        raise ValueError("Tracking workbook has no worksheets.")
    return wb.active


def _write_log_row(ws: Worksheet, row_idx: int, row_data: Tuple[Any, ...]) -> None:
    """Write a 10-column tracking row obeying ERR_XLSX_007 (@) and E13 (dynamic height)."""
    stt, dt_val, mod, code, name, c_type, rules, ui, status, notes = row_data
    specs = (
        ("A", stt, _FONT_DATA, _ALIGN_WRAP_CENTER, "@"),
        ("B", dt_val, _FONT_DATA, _ALIGN_WRAP_CENTER, "DD/MM/YYYY"),
        ("C", mod, _FONT_DATA, _ALIGN_WRAP_CENTER, "@"),
        ("D", code, _FONT_CODE, _ALIGN_WRAP_CENTER, "@"),
        ("E", name, _FONT_DATA_BOLD, _ALIGN_WRAP_LEFT, "@"),
        ("F", c_type, _FONT_DATA, _ALIGN_WRAP_CENTER, "@"),
        ("G", rules, _FONT_DATA, _ALIGN_WRAP_LEFT, "@"),
        ("H", ui, _FONT_DATA, _ALIGN_WRAP_LEFT, "@"),
        ("I", status, _FONT_DATA, _ALIGN_WRAP_CENTER, "@"),
        ("J", notes, _FONT_DATA, _ALIGN_WRAP_LEFT, "@"),
    )
    fill = _FILL_ZEBRA if row_idx % 2 == 0 else _FILL_WHITE
    max_lines = 1

    for col_letter, val, font_obj, align_obj, num_fmt in specs:
        cell = ws[f"{col_letter}{row_idx}"]
        cell.value = sanitize_xml_string(val) if isinstance(val, str) else val
        cell.font = font_obj
        cell.alignment = align_obj
        cell.fill = fill
        cell.border = _BORDER_THIN
        cell.number_format = num_fmt
        if isinstance(val, str):
            max_lines = max(max_lines, val.count("\n") + 1, len(val) // 44 + 1)

    ws.row_dimensions[row_idx].height = max(26.0, max_lines * 16.5)


def _build_overview_sheet(ws: Worksheet) -> None:
    """Populate Module_Overview with live COUNTIFS & SUM formulas (E3, E14)."""
    ws.views.sheetView[0].showGridLines = True
    ws.merge_cells("A1:H1")
    ws["A1"] = "TỔNG QUAN TIẾN ĐỘ CÔNG VIỆC & HỒ SƠ THEO BỘ PHẬN"
    ws["A1"].font = _FONT_TITLE
    ws["A1"].fill = _FILL_NAVY
    ws["A1"].alignment = _ALIGN_CENTER
    ws.row_dimensions[1].height = 36

    headers = [
        ("STT", 6),
        ("Module / Bộ phận", 26),
        ("Tổng đầu việc", 16),
        ("Đang soạn thảo", 16),
        ("Chờ Review", 16),
        ("Đã duyệt", 16),
        ("Đã chốt", 16),
        ("Tỉ lệ hoàn thành", 18),
    ]
    ws.row_dimensions[3].height = 28
    for idx, (label, width) in enumerate(headers, start=1):
        col_l = get_column_letter(idx)
        cell = ws[f"{col_l}3"]
        cell.value = label
        cell.font = _FONT_HDR_DARK
        cell.fill = _FILL_SUBHEADER
        cell.border = _BORDER_HEADER
        cell.alignment = _ALIGN_WRAP_CENTER
        ws.column_dimensions[col_l].width = width

    modules = [
        "Kế toán - Tài chính",
        "Hành chính - Nhân sự",
        "Vận hành & Kho",
        "Kinh doanh & CSKH",
        "Báo cáo Tổng hợp",
    ]
    start_r = 4
    for idx, mod in enumerate(modules, start=1):
        r = start_r + idx - 1
        ws[f"A{r}"] = idx
        ws[f"B{r}"] = mod
        ws[f"C{r}"] = f'=COUNTIFS(FSD_Mockup_Log!$C$3:$C$500, B{r})'
        ws[f"D{r}"] = f'=COUNTIFS(FSD_Mockup_Log!$C$3:$C$500, B{r}, FSD_Mockup_Log!$I$3:$I$500, "Đang soạn thảo")'
        ws[f"E{r}"] = f'=COUNTIFS(FSD_Mockup_Log!$C$3:$C$500, B{r}, FSD_Mockup_Log!$I$3:$I$500, "Chờ Review")'
        ws[f"F{r}"] = f'=COUNTIFS(FSD_Mockup_Log!$C$3:$C$500, B{r}, FSD_Mockup_Log!$I$3:$I$500, "Đã duyệt")'
        ws[f"G{r}"] = f'=COUNTIFS(FSD_Mockup_Log!$C$3:$C$500, B{r}, FSD_Mockup_Log!$I$3:$I$500, "Đã chốt")'
        ws[f"H{r}"] = f"=IF(C{r}=0, 0, (F{r}+G{r})/C{r})"

        for c_char in ("A", "B", "C", "D", "E", "F", "G", "H"):
            cell = ws[f"{c_char}{r}"]
            cell.font = _FONT_DATA
            cell.border = _BORDER_THIN
            cell.fill = _FILL_ZEBRA if idx % 2 == 0 else _FILL_WHITE
            cell.alignment = _ALIGN_LEFT if c_char == "B" else _ALIGN_CENTER
        ws[f"H{r}"].number_format = "0.0%"
        ws.row_dimensions[r].height = 22

    total_r = start_r + len(modules)
    ws[f"A{total_r}"] = ""
    ws[f"B{total_r}"] = "TỔNG CỘNG"
    for c_char in ("C", "D", "E", "F", "G"):
        ws[f"{c_char}{total_r}"] = f"=SUM({c_char}{start_r}:{c_char}{total_r - 1})"
    ws[f"H{total_r}"] = f"=IF(C{total_r}=0, 0, (F{total_r}+G{total_r})/C{total_r})"

    for c_char in ("A", "B", "C", "D", "E", "F", "G", "H"):
        cell = ws[f"{c_char}{total_r}"]
        cell.font = _FONT_DATA_BOLD
        cell.border = _BORDER_TOTAL
        cell.fill = _FILL_SUBHEADER
        cell.alignment = _ALIGN_LEFT if c_char == "B" else _ALIGN_CENTER
    ws[f"H{total_r}"].number_format = "0.0%"
    ws.row_dimensions[total_r].height = 26
    apply_worksheet_print_setup(ws, orientation="landscape", print_title_header_row=3)


class OfficeWorkTracker:
    """
    Unified Office Work Log & Daily Task Tracker.

    Operates seamlessly on both legacy `project_tracking.xlsx` / `task_tracking.xlsx`
    files and newly initialized Office workbooks.
    """

    def __init__(
        self,
        project_file: Path | str = DEFAULT_PROJECT_FILE,
        task_file: Path | str = DEFAULT_TASK_FILE,
    ) -> None:
        self.project_file = Path(project_file)
        self.task_file = Path(task_file)

    def init_project_tracking(self, force: bool = False) -> Dict[str, Any]:
        """Initialize the master work log workbook with Office IT sample records."""
        if self.project_file.exists() and not force:
            return {
                "ok": True,
                "created": False,
                "path": str(self.project_file),
                "message": "Workbook already exists (skipped overwrite).",
            }

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "FSD_Mockup_Log"
        ws.views.sheetView[0].showGridLines = True

        ws.merge_cells("A1:J1")
        ws["A1"] = "NHẬT KÝ THEO DÕI CÔNG VIỆC VĂN PHÒNG & XỬ LÝ TÀI LIỆU"
        ws["A1"].font = _FONT_TITLE
        ws["A1"].fill = _FILL_NAVY
        ws["A1"].alignment = _ALIGN_CENTER
        ws.row_dimensions[1].height = 36

        ws.row_dimensions[2].height = 30
        for idx, (label, width) in enumerate(_LOG_HEADERS, start=1):
            col_l = get_column_letter(idx)
            cell = ws[f"{col_l}2"]
            cell.value = label
            cell.font = _FONT_HDR
            cell.fill = _FILL_NAVY
            cell.border = _BORDER_HEADER
            cell.alignment = _ALIGN_WRAP_CENTER
            ws.column_dimensions[col_l].width = width

        ws.freeze_panes = "A3"
        ws.auto_filter.ref = "A2:J2"

        today_str = datetime.date.today().strftime("%d/%m/%Y")
        sample_rows = [
            (
                1,
                today_str,
                "Kế toán - Tài chính",
                "FIN-PAY-01",
                "Bảng lương Tổng hợp Tháng 11",
                "Tạo mới",
                "Chuẩn hóa công thức SUMIFS theo phòng ban, bảo toàn mã NV/CCCD dạng text '@'.",
                "Xuất báo cáo tổng hợp DOCX + chèn biểu đồ cột chi phí lương.",
                "Đã duyệt",
                "Khởi tạo mẫu chuẩn Office Suite",
            ),
            (
                2,
                today_str,
                "Hành chính - Nhân sự",
                "ADM-DOC-02",
                "Công văn Triển khai Quy trình Số",
                "Bổ sung",
                "Áp dụng thể thức Nghị định 30/2020/NĐ-CP và chèn dấu giáp lai PDF tự động.",
                "Trộn thư hàng loạt từ danh sách nhân sự Excel sang DOCX.",
                "Chờ Review",
                "Đang kiểm tra in thử PDF/A",
            ),
        ]
        for r_idx, row_data in enumerate(sample_rows, start=3):
            _write_log_row(ws, r_idx, row_data)

        apply_worksheet_print_setup(ws, orientation="landscape", print_title_header_row=2)

        ws_overview = wb.create_sheet(title="Module_Overview")
        _build_overview_sheet(ws_overview)

        self.project_file.parent.mkdir(parents=True, exist_ok=True)
        wb.save(str(self.project_file))
        return {
            "ok": True,
            "created": True,
            "path": str(self.project_file),
            "message": f"Initialized workbook with {len(sample_rows)} Office sample records.",
        }

    def add_entry(
        self,
        module: str,
        code: str,
        name: str,
        change_type: str,
        rules: str,
        ui: str = "",
        status: str = "Đang soạn thảo",
        notes: str = "",
        date_str: Optional[str] = None,
    ) -> Dict[str, Any]:
        """Append a new tracked work record to `FSD_Mockup_Log` with next STT."""
        if not self.project_file.exists():
            self.init_project_tracking(force=False)

        wb = openpyxl.load_workbook(str(self.project_file))
        ws = _resolve_target_sheet(wb)

        last_stt = 0
        max_row = ws.max_row or 2
        for r in range(3, max_row + 1):
            v = ws.cell(r, 1).value
            if v is not None and str(v).strip().isdigit():
                last_stt = max(last_stt, int(str(v).strip()))
            elif ws.cell(r, 4).value is not None or ws.cell(r, 5).value is not None:
                last_stt += 1

        next_stt = last_stt + 1
        target_row = max(3, max_row + 1)
        actual_date = (date_str or "").strip() or datetime.date.today().strftime("%d/%m/%Y")

        row_tuple = (
            next_stt,
            actual_date,
            module.strip(),
            code.strip(),
            name.strip(),
            change_type.strip(),
            rules.strip(),
            ui.strip(),
            status.strip(),
            notes.strip(),
        )
        _write_log_row(ws, target_row, row_tuple)
        wb.save(str(self.project_file))

        return {
            "ok": True,
            "stt": next_stt,
            "row": target_row,
            "code": code.strip(),
            "name": name.strip(),
            "date": actual_date,
            "path": str(self.project_file),
        }

    def list_entries(
        self,
        date_filter: Optional[str] = None,
        module_filter: Optional[str] = None,
        status_filter: Optional[str] = None,
    ) -> List[Dict[str, Any]]:
        """Read work log records with optional date/module/status filtering."""
        if not self.project_file.exists():
            return []

        wb = openpyxl.load_workbook(str(self.project_file), data_only=True)
        ws = _resolve_target_sheet(wb)
        records: List[Dict[str, Any]] = []

        for r in range(3, (ws.max_row or 2) + 1):
            stt = ws.cell(r, 1).value
            dt_val = ws.cell(r, 2).value
            mod = ws.cell(r, 3).value
            code = ws.cell(r, 4).value
            name = ws.cell(r, 5).value
            c_type = ws.cell(r, 6).value
            rules = ws.cell(r, 7).value
            ui = ws.cell(r, 8).value
            status = ws.cell(r, 9).value
            notes = ws.cell(r, 10).value

            if not any(v is not None and str(v).strip() != "" for v in (code, name, rules)):
                continue

            if isinstance(dt_val, (datetime.date, datetime.datetime)):
                dt_str = dt_val.strftime("%d/%m/%Y")
            else:
                dt_str = str(dt_val or "").strip()

            if date_filter and date_filter.strip() != dt_str:
                continue
            if module_filter and module_filter.strip().lower() not in str(mod or "").lower():
                continue
            if status_filter and status_filter.strip().lower() != str(status or "").strip().lower():
                continue

            records.append(
                {
                    "row": r,
                    "stt": stt,
                    "date": dt_str,
                    "module": str(mod or ""),
                    "code": str(code or ""),
                    "name": str(name or ""),
                    "type": str(c_type or ""),
                    "rules": str(rules or ""),
                    "ui": str(ui or ""),
                    "status": str(status or ""),
                    "notes": str(notes or ""),
                }
            )
        return records

    def sync_daily_report(
        self,
        date_str: Optional[str] = None,
        apply_to_task_tracking: bool = False,
    ) -> Dict[str, Any]:
        """Synthesize daily task summary and optionally inject rows into task_tracking.xlsx."""
        target_date = (date_str or "").strip() or datetime.date.today().strftime("%d/%m/%Y")
        entries = self.list_entries(date_filter=target_date)

        # Build concise task items
        generated_tasks: List[Dict[str, str]] = []
        for idx, item in enumerate(entries, start=1):
            action_verb = "Xử lý"
            if item["type"] == "Tạo mới":
                action_verb = "Tạo mới & hoàn thiện"
            elif item["type"] == "Bổ sung":
                action_verb = "Bổ sung nội dung"
            elif item["type"] == "Chỉnh sửa":
                action_verb = "Cập nhật & chỉnh sửa"
            elif item["type"] == "Hủy bỏ":
                action_verb = "Rà soát đóng"

            summary_desc = f"{action_verb} [{item['code']}] {item['name']}".strip()
            task_status = "Done" if item["status"] in ("Đã duyệt", "Đã chốt") else "In Progress"
            generated_tasks.append(
                {
                    "stt": str(idx),
                    "module": item["module"] or "General",
                    "description": summary_desc,
                    "status": task_status,
                    "notes": item["notes"] or item["rules"][:120],
                }
            )

        injected = False
        target_sheet = None
        if apply_to_task_tracking and generated_tasks:
            if not self.task_file.exists():
                raise FileNotFoundError(f"Task tracking workbook not found: {self.task_file}")
            day_month = "-".join(target_date.split("/")[:2]) if "/" in target_date else target_date[:5]
            wb_task = openpyxl.load_workbook(str(self.task_file))
            if day_month not in wb_task.sheetnames:
                raise ValueError(
                    f"Daily sheet '{day_month}' not found in {self.task_file.name}. "
                    f"Available: {wb_task.sheetnames}"
                )
            target_sheet = day_month
            ws_t = wb_task[day_month]

            insert_r = 3
            while insert_r <= (ws_t.max_row or 3) + 1:
                c_val = ws_t.cell(insert_r, 3).value
                if c_val is None or str(c_val).strip() == "":
                    break
                insert_r += 1

            for item in generated_tasks:
                ws_t.cell(insert_r, 1, value=insert_r - 2)
                ws_t.cell(insert_r, 2, value=item["module"])
                ws_t.cell(insert_r, 3, value=item["description"])
                ws_t.cell(insert_r, 4, value=item["status"])
                ws_t.cell(insert_r, 5, value=item["notes"])
                for col_i in range(1, 6):
                    c = ws_t.cell(insert_r, col_i)
                    c.font = _FONT_DATA
                    c.border = _BORDER_THIN
                    c.alignment = _ALIGN_WRAP_LEFT if col_i in (3, 5) else _ALIGN_WRAP_CENTER
                ws_t.row_dimensions[insert_r].height = 24
                insert_r += 1

            wb_task.save(str(self.task_file))
            injected = True

        return {
            "ok": True,
            "date": target_date,
            "entries_matched": len(entries),
            "tasks": generated_tasks,
            "injected_to_task_file": injected,
            "task_file": str(self.task_file),
            "target_sheet": target_sheet,
        }

    def inspect_workbooks(self) -> Dict[str, Any]:
        """Return lightweight status metadata of both tracking workbooks for the UI."""
        entries = self.list_entries()
        status_counts: Dict[str, int] = {}
        for item in entries:
            st = item["status"] or "Khác"
            status_counts[st] = status_counts.get(st, 0) + 1

        task_sheets: List[str] = []
        if self.task_file.exists():
            wb_t = openpyxl.load_workbook(str(self.task_file), read_only=True)
            task_sheets = list(wb_t.sheetnames)
            wb_t.close()

        return {
            "ok": True,
            "project_file": str(self.project_file),
            "project_exists": self.project_file.exists(),
            "total_entries": len(entries),
            "status_counts": status_counts,
            "recent_entries": entries[-15:][::-1],
            "task_file": str(self.task_file),
            "task_exists": self.task_file.exists(),
            "task_sheets": task_sheets,
        }

