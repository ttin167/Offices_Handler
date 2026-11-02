"""
core.ba_hub.testcase_generator — Production UAT Test Case & Acceptance Matrix Generator.

Translates:
- Gherkin Scenarios (Given-When-Then)
- Use Case Main & Exception flows
into publication-grade, interactive Excel (.xlsx) UAT Checklists with dropdown validations.
"""

from __future__ import annotations

import datetime
import os
import re
from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.datavalidation import DataValidation


@dataclass
class UATScenario:
    test_id: str
    req_ref: str
    title: str
    test_type: str = "Functional"
    preconditions: str = ""
    steps: str = ""
    test_data: str = ""
    expected_result: str = ""
    actual_result: str = ""
    status: str = "NOT_RUN"  # PASS, FAIL, BLOCKED, NOT_RUN
    severity: str = "HIGH"   # CRITICAL, HIGH, MEDIUM, LOW
    notes: str = ""

    def to_dict(self) -> Dict[str, Any]:
        return {
            "test_id": self.test_id,
            "req_ref": self.req_ref,
            "title": self.title,
            "test_type": self.test_type,
            "preconditions": self.preconditions,
            "steps": self.steps,
            "test_data": self.test_data,
            "expected_result": self.expected_result,
            "actual_result": self.actual_result,
            "status": self.status,
            "severity": self.severity,
            "notes": self.notes,
        }


class UATTestcaseGenerator:
    """Translates User Stories and Use Cases into UAT Checklists."""

    @classmethod
    def generate_from_stories(cls, stories_data: List[Dict[str, Any]]) -> List[UATScenario]:
        test_cases: List[UATScenario] = []
        tc_counter = 1

        for st in stories_data:
            req_id = st.get("id", "REQ-001")
            scenarios = st.get("scenarios", [])

            for sc in scenarios:
                test_id = f"TC-UAT-{tc_counter:03d}"
                tc_counter += 1

                title = sc.get("title", "Kịch bản kiểm thử")
                sc_type = sc.get("type", "Functional")
                given = sc.get("given", "")
                when = sc.get("when", "")
                then = sc.get("then", "")
                and_cond = sc.get("and", "")

                step_text = f"1. Khởi động môi trường kiểm thử.\n2. Thực hiện hành động: {when}."
                exp_text = f"Kết quả mong đợi:\n- {then}"
                if and_cond:
                    exp_text += f"\n- {and_cond}"

                test_cases.append(UATScenario(
                    test_id=test_id,
                    req_ref=req_id,
                    title=title,
                    test_type=sc_type,
                    preconditions=given,
                    steps=step_text,
                    test_data="Dữ liệu mẫu kiểm thử theo kịch bản",
                    expected_result=exp_text,
                    actual_result="",
                    status="NOT_RUN",
                    severity="CRITICAL" if "happy" in sc_type.lower() else "HIGH",
                    notes=f"Được sinh từ User Story {req_id}: '{st.get('title', '')}'"
                ))

        return test_cases

    @classmethod
    def export_uat_checklist_to_excel(
        cls,
        project_name: str,
        test_cases: List[UATScenario],
        output_path: str,
        module_name: str = "Phân hệ Nghiệp vụ"
    ) -> str:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "UAT_Acceptance_Checklist"
        ws.views.sheetView[0].showGridLines = True

        # Document Header
        ws.merge_cells("A1:K1")
        title_cell = ws["A1"]
        title_cell.value = f"KỊCH BẢN KIỂM THỬ NGHIỆM THU NGƯỜI DÙNG (UAT CHECKLIST) — DỰ ÁN: {project_name.upper()}"
        title_cell.font = Font(name="Segoe UI", size=14, bold=True, color="FFFFFF")
        title_cell.fill = PatternFill(start_color="0F172A", end_color="0F172A", fill_type="solid")
        title_cell.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[1].height = 40

        ws.merge_cells("A2:K2")
        sub_cell = ws["A2"]
        sub_cell.value = f"Phân hệ: {module_name} | Tiêu chuẩn nghiệm thu chức năng BABOK / Agile UAT | Ngày lập: {datetime.date.today().strftime('%d/%m/%Y')}"
        sub_cell.font = Font(name="Segoe UI", size=10, italic=True, color="64748B")
        sub_cell.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[2].height = 22

        headers = [
            "Mã Test (Test ID)",
            "Mã Yêu Cầu (Req ID)",
            "Tên Kịch Bản Nghiệm Thu (Scenario)",
            "Loại Kịch Bản",
            "Tiền Điều Kiện (Preconditions)",
            "Các Bước Thực Hiện (Test Steps)",
            "Kết Quả Mong Đợi (Expected Outcome)",
            "Kết Quả Thực Tế (Actual Outcome)",
            "Trạng Thái (Status)",
            "Mức Độ (Severity)",
            "Ghi Chú & Ký Nhận (Sign-off)"
        ]

        header_row = 4
        ws.row_dimensions[header_row].height = 30
        h_fill = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")
        h_font = Font(name="Segoe UI", size=10, bold=True, color="FFFFFF")
        h_border = Border(
            top=Side(style="thin", color="CBD5E1"), bottom=Side(style="medium", color="0284C7"),
            left=Side(style="thin", color="CBD5E1"), right=Side(style="thin", color="CBD5E1")
        )

        for col_idx, h in enumerate(headers, start=1):
            cell = ws.cell(row=header_row, column=col_idx, value=h)
            cell.fill = h_fill
            cell.font = h_font
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = h_border

        # Data rows
        cell_border = Border(
            top=Side(style="thin", color="E2E8F0"), bottom=Side(style="thin", color="E2E8F0"),
            left=Side(style="thin", color="E2E8F0"), right=Side(style="thin", color="E2E8F0")
        )
        zebra_fill = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")

        start_row = 5
        for idx, tc in enumerate(test_cases, start=1):
            curr_r = start_row + idx - 1
            ws.row_dimensions[curr_r].height = 55  # Multiline comfortable height

            ws.cell(row=curr_r, column=1, value=tc.test_id).font = Font(name="Consolas", size=9.5, bold=True, color="0369A1")
            ws.cell(row=curr_r, column=1).alignment = Alignment(horizontal="center", vertical="center")

            ws.cell(row=curr_r, column=2, value=tc.req_ref).font = Font(name="Consolas", size=9.5, color="475569")
            ws.cell(row=curr_r, column=2).alignment = Alignment(horizontal="center", vertical="center")

            ws.cell(row=curr_r, column=3, value=tc.title).font = Font(name="Segoe UI", size=9.5, bold=True)
            ws.cell(row=curr_r, column=3).alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

            ws.cell(row=curr_r, column=4, value=tc.test_type).alignment = Alignment(horizontal="center", vertical="center")
            ws.cell(row=curr_r, column=5, value=tc.preconditions).alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
            ws.cell(row=curr_r, column=6, value=tc.steps).alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
            ws.cell(row=curr_r, column=7, value=tc.expected_result).alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
            ws.cell(row=curr_r, column=8, value=tc.actual_result).alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

            c_status = ws.cell(row=curr_r, column=9, value=tc.status)
            c_status.alignment = Alignment(horizontal="center", vertical="center")
            c_status.font = Font(name="Segoe UI", size=9.5, bold=True, color="0F172A")

            c_sev = ws.cell(row=curr_r, column=10, value=tc.severity)
            c_sev.alignment = Alignment(horizontal="center", vertical="center")
            c_sev.font = Font(name="Segoe UI", size=9.5, bold=True, color="BE123C" if tc.severity == "CRITICAL" else "D97706")

            ws.cell(row=curr_r, column=11, value=tc.notes).alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

            for col in range(1, 12):
                c = ws.cell(row=curr_r, column=col)
                c.border = cell_border
                if not c.font or c.font.name != "Consolas":
                    c.font = Font(name="Segoe UI", size=9)
                if idx % 2 == 0 and c.fill.fill_type is None:
                    c.fill = zebra_fill

        # Data Validation Dropdowns for Status and Severity
        max_row = start_row + len(test_cases)
        dv_status = DataValidation(type="list", formula1='"PASS,FAIL,BLOCKED,NOT_RUN"', allow_blank=True)
        dv_status.error = "Vui lòng chọn trạng thái từ danh sách (PASS, FAIL, BLOCKED, NOT_RUN)"
        dv_status.errorTitle = "Trạng thái không hợp lệ"
        ws.add_data_validation(dv_status)
        dv_status.add(f"I5:I{max_row}")

        dv_sev = DataValidation(type="list", formula1='"CRITICAL,HIGH,MEDIUM,LOW"', allow_blank=True)
        dv_sev.error = "Vui lòng chọn mức độ nghiêm trọng từ danh sách"
        dv_sev.errorTitle = "Mức độ không hợp lệ"
        ws.add_data_validation(dv_sev)
        dv_sev.add(f"J5:J{max_row}")

        # Column widths
        ws.column_dimensions["A"].width = 14
        ws.column_dimensions["B"].width = 14
        ws.column_dimensions["C"].width = 28
        ws.column_dimensions["D"].width = 15
        ws.column_dimensions["E"].width = 26
        ws.column_dimensions["F"].width = 30
        ws.column_dimensions["G"].width = 32
        ws.column_dimensions["H"].width = 25
        ws.column_dimensions["I"].width = 15
        ws.column_dimensions["J"].width = 14
        ws.column_dimensions["K"].width = 25

        ws.freeze_panes = "A5"

        os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
        wb.save(output_path)
        return output_path

