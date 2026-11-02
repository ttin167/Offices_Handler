"""
core.ba_hub.reconciliation — Two-Way Enterprise Data Reconciliation Engine for BAs.

Crucial for:
1. Reconciling Core Banking vs Payment Gateway transactions.
2. Warehouse Stock Records (WMS) vs Accounting Ledger (ERP).
3. Data Migration validation (Legacy DB vs Target DB).
4. Automated KPI Dashboard & Multi-Sheet Excel Variance Reports.
"""

from __future__ import annotations

import datetime
import os
import re
from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional, Set, Tuple

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


@dataclass
class DiscrepancyDetail:
    key: str
    column: str
    val_a: Any
    val_b: Any
    diff_numeric: Optional[float] = None
    row_data_a: Dict[str, Any] = field(default_factory=dict)
    row_data_b: Dict[str, Any] = field(default_factory=dict)


@dataclass
class ReconciliationSummary:
    source_a_name: str
    source_b_name: str
    total_a: int
    total_b: int
    matched_exact: int
    matched_discrepant: int
    missing_in_b: int
    missing_in_a: int
    match_rate_pct: float
    total_variance_sum: float
    key_columns: List[str]
    discrepancies: List[DiscrepancyDetail] = field(default_factory=list)
    missing_b_records: List[Dict[str, Any]] = field(default_factory=list)
    missing_a_records: List[Dict[str, Any]] = field(default_factory=list)


class DataReconciliationEngine:
    """Performs deep row-level and column-level reconciliation between two tabular datasets."""

    @staticmethod
    def _read_sheet_data(file_path: str, sheet_name: Optional[str] = None) -> Tuple[List[str], List[Dict[str, Any]]]:
        wb = openpyxl.load_workbook(file_path, data_only=True)
        ws = wb[sheet_name] if sheet_name and sheet_name in wb.sheetnames else wb.active
        rows = list(ws.iter_rows(values_only=True))
        if not rows:
            return [], []

        # Find header
        header_idx = 0
        for idx, row in enumerate(rows[:5]):
            str_cells = [c for c in row if c is not None and str(c).strip()]
            if len(str_cells) >= 2:
                header_idx = idx
                break

        raw_headers = [str(c).strip() if c is not None else f"Col_{i}" for i, c in enumerate(rows[header_idx])]
        records: List[Dict[str, Any]] = []

        for r in rows[header_idx + 1:]:
            if all(c is None or str(c).strip() == "" for c in r):
                continue
            rec = {raw_headers[i]: r[i] for i in range(min(len(raw_headers), len(r)))}
            records.append(rec)

        return raw_headers, records

    @classmethod
    def reconcile(
        cls,
        source_a: Any,  # file_path str OR List[Dict[str, Any]]
        source_b: Any,
        key_columns: List[str],
        compare_columns: Optional[List[str]] = None,
        numeric_tolerance: float = 0.0,
        source_a_name: str = "Hệ thống A (Sổ kế toán / Core)",
        source_b_name: str = "Hệ thống B (Đối tác / Thực tế)",
    ) -> ReconciliationSummary:
        # Load records
        if isinstance(source_a, str):
            headers_a, records_a = cls._read_sheet_data(source_a)
        else:
            records_a = list(source_a)
            headers_a = list(records_a[0].keys()) if records_a else []

        if isinstance(source_b, str):
            headers_b, records_b = cls._read_sheet_data(source_b)
        else:
            records_b = list(source_b)
            headers_b = list(records_b[0].keys()) if records_b else []

        # Resolve compare columns
        if not compare_columns:
            common = [c for c in headers_a if c in headers_b and c not in key_columns]
            compare_columns = common or [c for c in headers_a if c not in key_columns]

        # Normalize key building
        def build_key(rec: Dict[str, Any]) -> str:
            parts = [str(rec.get(k, "")).strip().lower() for k in key_columns]
            return "||".join(parts)

        map_a = {build_key(r): r for r in records_a}
        map_b = {build_key(r): r for r in records_b}

        keys_a = set(map_a.keys())
        keys_b = set(map_b.keys())

        common_keys = keys_a.intersection(keys_b)
        missing_b_keys = keys_a - keys_b
        missing_a_keys = keys_b - keys_a

        matched_exact_count = 0
        discrepancies: List[DiscrepancyDetail] = []
        total_variance_sum = 0.0

        for key in common_keys:
            rec_a = map_a[key]
            rec_b = map_b[key]
            has_diff = False

            for col in compare_columns:
                val_a = rec_a.get(col)
                val_b = rec_b.get(col)

                # Check numeric equivalence
                num_a, num_b = None, None
                try:
                    if val_a is not None and str(val_a).strip() != "":
                        num_a = float(str(val_a).replace(",", ".").replace(" ", ""))
                    if val_b is not None and str(val_b).strip() != "":
                        num_b = float(str(val_b).replace(",", ".").replace(" ", ""))
                except ValueError:
                    pass

                if num_a is not None and num_b is not None:
                    diff = abs(num_a - num_b)
                    if diff > numeric_tolerance:
                        has_diff = True
                        total_variance_sum += diff
                        discrepancies.append(DiscrepancyDetail(
                            key=key, column=col, val_a=val_a, val_b=val_b, diff_numeric=diff,
                            row_data_a=rec_a, row_data_b=rec_b
                        ))
                else:
                    s_a = str(val_a).strip() if val_a is not None else ""
                    s_b = str(val_b).strip() if val_b is not None else ""
                    if s_a != s_b:
                        has_diff = True
                        discrepancies.append(DiscrepancyDetail(
                            key=key, column=col, val_a=val_a, val_b=val_b, diff_numeric=None,
                            row_data_a=rec_a, row_data_b=rec_b
                        ))

            if not has_diff:
                matched_exact_count += 1

        total_a = len(records_a)
        total_b = len(records_b)
        denom = max(total_a, total_b, 1)
        match_rate = round((matched_exact_count / denom) * 100, 2)

        return ReconciliationSummary(
            source_a_name=source_a_name,
            source_b_name=source_b_name,
            total_a=total_a,
            total_b=total_b,
            matched_exact=matched_exact_count,
            matched_discrepant=len({d.key for d in discrepancies}),
            missing_in_b=len(missing_b_keys),
            missing_in_a=len(missing_a_keys),
            match_rate_pct=match_rate,
            total_variance_sum=round(total_variance_sum, 2),
            key_columns=key_columns,
            discrepancies=discrepancies,
            missing_b_records=[map_a[k] for k in missing_b_keys],
            missing_a_records=[map_b[k] for k in missing_a_keys]
        )

    @classmethod
    def export_report_to_excel(cls, summary: ReconciliationSummary, output_path: str) -> str:
        wb = openpyxl.Workbook()

        # Styles
        font_header = Font(name="Segoe UI", size=10, bold=True, color="FFFFFF")
        fill_header = PatternFill(start_color="0F172A", end_color="0F172A", fill_type="solid")
        fill_zebra = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
        fill_diff = PatternFill(start_color="FEE2E2", end_color="FEE2E2", fill_type="solid")  # Red light
        font_diff = Font(name="Segoe UI", size=9.5, bold=True, color="991B1B")
        border_thin = Border(
            top=Side(style="thin", color="CBD5E1"), bottom=Side(style="thin", color="CBD5E1"),
            left=Side(style="thin", color="CBD5E1"), right=Side(style="thin", color="CBD5E1")
        )

        # ----------------------------------------------------------------------
        # SHEET 1: Summary Dashboard
        # ----------------------------------------------------------------------
        ws1 = wb.active
        ws1.title = "Recon_Executive_Summary"
        ws1.views.sheetView[0].showGridLines = True

        ws1.merge_cells("A1:F1")
        t_cell = ws1["A1"]
        t_cell.value = "BÁO CÁO ĐỐI SOÁT DỮ LIỆU & PHÂN TÍCH CHÊNH LỆCH (RECONCILIATION)"
        t_cell.font = Font(name="Segoe UI", size=14, bold=True, color="FFFFFF")
        t_cell.fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
        t_cell.alignment = Alignment(horizontal="center", vertical="center")
        ws1.row_dimensions[1].height = 40

        ws1.merge_cells("A2:F2")
        s_cell = ws1["A2"]
        s_cell.value = f"Đối chiếu: {summary.source_a_name} ↔ {summary.source_b_name} | Ngày chạy: {datetime.date.today().strftime('%d/%m/%Y')}"
        s_cell.font = Font(name="Segoe UI", size=10, italic=True, color="64748B")
        s_cell.alignment = Alignment(horizontal="center", vertical="center")
        ws1.row_dimensions[2].height = 22

        kpis = [
            ("Tỷ lệ khớp hoàn toàn (Match Rate)", f"{summary.match_rate_pct}%", "0284C7"),
            ("Tổng số dòng khớp 100% (Exact Matches)", summary.matched_exact, "16A34A"),
            ("Số dòng khớp Key nhưng Lệch giá trị (Discrepancies)", summary.matched_discrepant, "DC2626"),
            ("Số dòng có ở A nhưng thiếu ở B (Missing in B)", summary.missing_in_b, "D97706"),
            ("Số dòng có ở B nhưng thiếu ở A (Missing in A)", summary.missing_in_a, "7C3AED"),
            ("Tổng giá trị chênh lệch tuyệt đối (Total Variance)", f"{summary.total_variance_sum:,.2f}", "BE123C"),
        ]

        row_kpi = 4
        for title, val, color_hex in kpis:
            ws1.cell(row=row_kpi, column=2, value=title).font = Font(name="Segoe UI", size=10, bold=True, color="334155")
            c_v = ws1.cell(row=row_kpi, column=4, value=val)
            c_v.font = Font(name="Segoe UI", size=11, bold=True, color=color_hex)
            c_v.alignment = Alignment(horizontal="right", vertical="center")
            ws1.row_dimensions[row_kpi].height = 24
            for c in [2, 3, 4]:
                ws1.cell(row=row_kpi, column=c).border = border_thin
            row_kpi += 1

        ws1.column_dimensions["B"].width = 50
        ws1.column_dimensions["C"].width = 5
        ws1.column_dimensions["D"].width = 25

        # ----------------------------------------------------------------------
        # SHEET 2: Discrepancies Details
        # ----------------------------------------------------------------------
        ws2 = wb.create_sheet(title="Matched_Discrepancies")
        ws2.views.sheetView[0].showGridLines = True
        ws2.row_dimensions[1].height = 28

        h2 = ["STT", "Mã Khóa (Composite Key)", "Trường Dữ Liệu Bị Lệch", f"Giá Trị {summary.source_a_name}", f"Giá Trị {summary.source_b_name}", "Chênh Lệch Tuyệt Đối"]
        for idx, h in enumerate(h2, start=1):
            c = ws2.cell(row=1, column=idx, value=h)
            c.font = font_header
            c.fill = fill_header
            c.alignment = Alignment(horizontal="center", vertical="center")

        for idx, d in enumerate(summary.discrepancies, start=1):
            r = idx + 1
            ws2.row_dimensions[r].height = 22
            ws2.cell(row=r, column=1, value=idx).alignment = Alignment(horizontal="center", vertical="center")
            ws2.cell(row=r, column=2, value=d.key).font = Font(name="Consolas", size=9.5, bold=True, color="0369A1")
            ws2.cell(row=r, column=3, value=d.column).font = Font(name="Segoe UI", size=9.5, bold=True)
            c_a = ws2.cell(row=r, column=4, value=str(d.val_a))
            c_b = ws2.cell(row=r, column=5, value=str(d.val_b))
            for cell_item in [c_a, c_b]:
                cell_item.fill = fill_diff
                cell_item.font = font_diff
                cell_item.alignment = Alignment(horizontal="center", vertical="center")

            diff_str = f"{d.diff_numeric:,.2f}" if d.diff_numeric is not None else "Text Mismatch"
            ws2.cell(row=r, column=6, value=diff_str).alignment = Alignment(horizontal="right", vertical="center")

            for col in range(1, 7):
                ws2.cell(row=r, column=col).border = border_thin

        ws2.column_dimensions["A"].width = 6
        ws2.column_dimensions["B"].width = 25
        ws2.column_dimensions["C"].width = 25
        ws2.column_dimensions["D"].width = 25
        ws2.column_dimensions["E"].width = 25
        ws2.column_dimensions["F"].width = 22
        ws2.freeze_panes = "A2"

        # ----------------------------------------------------------------------
        # SHEET 3: Missing in B
        # ----------------------------------------------------------------------
        ws3 = wb.create_sheet(title=f"Missing_In_{summary.source_b_name[:15]}")
        ws3.views.sheetView[0].showGridLines = True
        if summary.missing_b_records:
            cols3 = list(summary.missing_b_records[0].keys())
            for idx, c in enumerate(cols3, start=1):
                cell = ws3.cell(row=1, column=idx, value=c)
                cell.font = font_header
                cell.fill = PatternFill(start_color="D97706", end_color="D97706", fill_type="solid")  # Amber
            for r_idx, rec in enumerate(summary.missing_b_records, start=2):
                for c_idx, col_name in enumerate(cols3, start=1):
                    ws3.cell(row=r_idx, column=c_idx, value=rec.get(col_name)).border = border_thin

        # ----------------------------------------------------------------------
        # SHEET 4: Missing in A
        # ----------------------------------------------------------------------
        ws4 = wb.create_sheet(title=f"Missing_In_{summary.source_a_name[:15]}")
        ws4.views.sheetView[0].showGridLines = True
        if summary.missing_a_records:
            cols4 = list(summary.missing_a_records[0].keys())
            for idx, c in enumerate(cols4, start=1):
                cell = ws4.cell(row=1, column=idx, value=c)
                cell.font = font_header
                cell.fill = PatternFill(start_color="7C3AED", end_color="7C3AED", fill_type="solid")  # Purple
            for r_idx, rec in enumerate(summary.missing_a_records, start=2):
                for c_idx, col_name in enumerate(cols4, start=1):
                    ws4.cell(row=r_idx, column=c_idx, value=rec.get(col_name)).border = border_thin

        os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
        wb.save(output_path)
        return output_path

