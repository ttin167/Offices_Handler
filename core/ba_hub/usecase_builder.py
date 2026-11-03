"""
core.ba_hub.usecase_builder — Production Use Case Specification Engine for Business Analysts.

Generates complete, publication-grade Use Case Specifications with:
1. Actor-System interaction clarity.
2. Main Success Flow, Alternative Flows, and Exception Flows.
3. OpenXML Word Table formatting with cantSplit, tblHeader, and professional Slate/Navy typography.
4. BABOK Use Case Quality Audit Linter.
"""

from __future__ import annotations

import os
from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional

import docx
from docx import Document
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
from docx.shared import Inches, Pt, RGBColor


@dataclass
class FlowStep:
    step: str
    actor: str
    action: str

    def to_dict(self) -> Dict[str, str]:
        return {"step": self.step, "actor": self.actor, "action": self.action}


@dataclass
class BranchFlow:
    id: str
    condition: str
    steps: List[str] = field(default_factory=list)

    def to_dict(self) -> Dict[str, Any]:
        return {"id": self.id, "condition": self.condition, "steps": self.steps}


@dataclass
class UseCaseSpec:
    id: str
    name: str
    module: str
    primary_actor: str
    secondary_actors: List[str] = field(default_factory=list)
    preconditions: List[str] = field(default_factory=list)
    postconditions: List[str] = field(default_factory=list)
    trigger: str = ""
    main_flow: List[FlowStep] = field(default_factory=list)
    alternative_flows: List[BranchFlow] = field(default_factory=list)
    exception_flows: List[BranchFlow] = field(default_factory=list)
    business_rules: List[str] = field(default_factory=list)

    def to_dict(self) -> Dict[str, Any]:
        return {
            "id": self.id,
            "name": self.name,
            "module": self.module,
            "primary_actor": self.primary_actor,
            "secondary_actors": self.secondary_actors,
            "preconditions": self.preconditions,
            "postconditions": self.postconditions,
            "trigger": self.trigger,
            "main_flow": [s.to_dict() for s in self.main_flow],
            "alternative_flows": [f.to_dict() for f in self.alternative_flows],
            "exception_flows": [f.to_dict() for f in self.exception_flows],
            "business_rules": self.business_rules,
        }

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> UseCaseSpec:
        return UseCaseBuilder.parse_dict(data)

    @property
    def title(self) -> str:
        return self.name

    @property
    def actor_primary(self) -> str:
        return self.primary_actor

    @property
    def pre_conditions(self) -> List[str]:
        return self.preconditions

    @property
    def post_conditions(self) -> List[str]:
        return self.postconditions

    @property
    def alt_flows(self) -> List[BranchFlow]:
        return self.alternative_flows

    @property
    def status(self) -> str:
        return "Approved"


class UseCaseQualityLinter:
    """Audits Use Case specification quality against BABOK & Cockburn guidelines."""

    @classmethod
    def audit_usecase(cls, data: Any) -> Dict[str, Any]:
        if hasattr(data, "to_dict"):
            data = data.to_dict()
        elif not isinstance(data, dict):
            data = {}
        score = 100
        findings: List[Dict[str, str]] = []

        # 1. Identification check
        if not data.get("id"):
            score -= 10
            findings.append({"severity": "WARNING", "message": "Thiếu mã định danh Use Case ID (ví dụ: UC-AUTH-001)."})
        if not data.get("name"):
            score -= 20
            findings.append({"severity": "CRITICAL", "message": "Thiếu tên Use Case. Tên nên bắt đầu bằng động từ hành động (ví dụ: 'Chuyển tiền nhanh', 'Đăng ký tài khoản')."})

        # 2. Actors & Trigger
        if not data.get("primary_actor"):
            score -= 15
            findings.append({"severity": "CRITICAL", "message": "Chưa chỉ định Tác nhân chính (Primary Actor)."})
        if not data.get("trigger"):
            score -= 10
            findings.append({"severity": "WARNING", "message": "Chưa mô tả sự kiện kích hoạt (Trigger Event) làm khởi đầu Use Case."})

        # 3. Pre & Post Conditions
        if not data.get("preconditions"):
            score -= 10
            findings.append({"severity": "WARNING", "message": "Thiếu tiền điều kiện (Preconditions) cần thỏa mãn trước khi thực thi."})
        if not data.get("postconditions"):
            score -= 10
            findings.append({"severity": "WARNING", "message": "Thiếu hậu điều kiện (Postconditions) sau khi Use Case kết thúc thành công."})

        # 4. Main Flow Steps
        main_flow = data.get("main_flow", [])
        if not main_flow:
            score -= 30
            findings.append({"severity": "CRITICAL", "message": "Chưa có các bước trong Luồng sự kiện chính (Main Success Flow)."})
        else:
            has_system = any("hệ thống" in str(s.get("actor", "")).lower() or "system" in str(s.get("actor", "")).lower() for s in main_flow)
            has_user = any("hệ thống" not in str(s.get("actor", "")).lower() and "system" not in str(s.get("actor", "")).lower() for s in main_flow)
            if not (has_system and has_user):
                score -= 10
                findings.append({"severity": "INFO", "message": "Luồng chính nên có sự tương tác 2 chiều rõ ràng giữa Tác nhân (Actor) và Hệ thống (System)."})

        # 5. Exception Flows
        if not data.get("exception_flows"):
            score -= 15
            findings.append({"severity": "WARNING", "message": "Chưa định nghĩa Luồng ngoại lệ (Exception Flows). Khi hệ thống gặp sự cố, chưa rõ phương án xử lý."})

        final_score = max(0, min(100, score))
        grade = "Xuất sắc (A+)" if final_score >= 90 else "Tốt (B)" if final_score >= 75 else "Cần bổ sung (C)" if final_score >= 50 else "Không đạt (D)"

        class AuditFinding(dict):
            def __getattr__(self, name):
                try:
                    return self[name]
                except KeyError:
                    raise AttributeError(name)

        class AuditResult(dict):
            def __getattr__(self, name):
                try:
                    return self[name]
                except KeyError:
                    raise AttributeError(name)

        return AuditResult({
            "score": final_score,
            "grade": grade,
            "is_valid": final_score >= 60,
            "findings": [AuditFinding(f) for f in findings],
            "steps_count": len(main_flow)
        })

    @classmethod
    def lint_usecase(cls, data: Any) -> Any:
        return cls.audit_usecase(data)


class UseCaseBuilder:
    """Builds and exports Use Case Specifications to Word DOCX and Markdown."""

    @staticmethod
    def parse_dict(data: Dict[str, Any]) -> UseCaseSpec:
        main_steps = [
            FlowStep(step=s.get("step", f"{idx}.0"), actor=s.get("actor", "Actor"), action=s.get("action", ""))
            for idx, s in enumerate(data.get("main_flow", []), start=1)
        ]
        alt_flows = [
            BranchFlow(id=f.get("id", f"Alt-{idx}"), condition=f.get("condition", ""), steps=f.get("steps", []))
            for idx, f in enumerate(data.get("alternative_flows", []), start=1)
        ]
        exc_flows = [
            BranchFlow(id=f.get("id", f"Exc-{idx}"), condition=f.get("condition", ""), steps=f.get("steps", []))
            for idx, f in enumerate(data.get("exception_flows", []), start=1)
        ]

        return UseCaseSpec(
            id=data.get("id", "UC-001"),
            name=data.get("name", "Đặc tả chức năng"),
            module=data.get("module", "Chung"),
            primary_actor=data.get("primary_actor", "Người dùng"),
            secondary_actors=data.get("secondary_actors", []),
            preconditions=data.get("preconditions", []),
            postconditions=data.get("postconditions", []),
            trigger=data.get("trigger", ""),
            main_flow=main_steps,
            alternative_flows=alt_flows,
            exception_flows=exc_flows,
            business_rules=data.get("business_rules", []),
        )

    @classmethod
    def export_to_docx(cls, uc: UseCaseSpec, output_path: str) -> str:
        doc = Document()
        for sec in doc.sections:
            sec.top_margin = Inches(0.8)
            sec.bottom_margin = Inches(0.8)
            sec.left_margin = Inches(0.9)
            sec.right_margin = Inches(0.9)

        # Title
        p_title = doc.add_paragraph()
        p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r_t = p_title.add_run(f"ĐẶC TẢ USE CASE: {uc.id} — {uc.name.upper()}")
        r_t.bold = True
        r_t.font.size = Pt(16)
        r_t.font.color.rgb = RGBColor(15, 23, 42)

        p_mod = doc.add_paragraph()
        p_mod.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r_m = p_mod.add_run(f"Phân hệ: {uc.module} | Chuẩn Tài Liệu Nghiệp Vụ Doanh Nghiệp")
        r_m.italic = True
        r_m.font.size = Pt(10)
        r_m.font.color.rgb = RGBColor(100, 116, 139)

        doc.add_paragraph().paragraph_format.space_after = Pt(10)

        # Build Specification Table
        table = doc.add_table(rows=0, cols=2)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        table.autofit = False

        col_w_label = Inches(1.8)
        col_w_val = Inches(4.8)

        def add_spec_row(label: str, content: Any, is_header: bool = False, bg_color: str = "FFFFFF"):
            row = table.add_row()
            trPr = row._tr.get_or_add_trPr()
            trPr.append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))
            if is_header:
                trPr.append(parse_xml(f'<w:tblHeader {nsdecls("w")}/>'))

            c0, c1 = row.cells[0], row.cells[1]
            c0.width, c1.width = col_w_label, col_w_val

            for cell, bg in [(c0, "F1F5F9" if not is_header else "1E293B"), (c1, bg_color)]:
                c_pr = cell._tc.get_or_add_tcPr()
                c_pr.append(parse_xml(f'<w:shd {nsdecls("w")} w:fill="{bg}"/>'))
                c_pr.append(parse_xml(f'<w:vAlign {nsdecls("w")} w:val="center"/>'))
                # Borders
                c_pr.append(parse_xml(f'<w:tcBorders {nsdecls("w")}><w:top w:val="single" w:sz="4" w:color="CBD5E1"/><w:bottom w:val="single" w:sz="4" w:color="CBD5E1"/><w:left w:val="single" w:sz="4" w:color="CBD5E1"/><w:right w:val="single" w:sz="4" w:color="CBD5E1"/></w:tcBorders>'))

            p0 = c0.paragraphs[0]
            r0 = p0.add_run(label)
            r0.bold = True
            r0.font.size = Pt(9.5)
            r0.font.color.rgb = RGBColor(255, 255, 255) if is_header else RGBColor(15, 23, 42)

            p1 = c1.paragraphs[0]
            if isinstance(content, str):
                r1 = p1.add_run(content)
                r1.font.size = Pt(9.5)
                if is_header:
                    r1.bold = True
                    r1.font.color.rgb = RGBColor(255, 255, 255)
            elif isinstance(content, list):
                for idx, item in enumerate(content):
                    if idx > 0:
                        p_sub = c1.add_paragraph()
                    else:
                        p_sub = p1
                    r_sub = p_sub.add_run(f"• {item}" if not isinstance(item, dict) else f"{item}")
                    r_sub.font.size = Pt(9.5)

        # Meta rows
        add_spec_row("Thuộc Tính", "Nội Dung Đặc Tả Chi Tiết", is_header=True)
        add_spec_row("Mã Use Case (ID)", uc.id)
        add_spec_row("Tên Use Case", uc.name)
        add_spec_row("Tác nhân chính", uc.primary_actor)
        add_spec_row("Tác nhân phụ / Hệ thống", ", ".join(uc.secondary_actors) if uc.secondary_actors else "Không có")
        add_spec_row("Sự kiện kích hoạt (Trigger)", uc.trigger or "Người dùng chọn chức năng")
        add_spec_row("Tiền điều kiện (Preconditions)", uc.preconditions if uc.preconditions else "Không có tiền điều kiện đặc biệt")
        add_spec_row("Hậu điều kiện (Postconditions)", uc.postconditions if uc.postconditions else "Dữ liệu được cập nhật thành công")

        # Main Flow Section
        row_mf_hdr = table.add_row()
        tr_mf = row_mf_hdr._tr.get_or_add_trPr()
        tr_mf.append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))
        c_full = row_mf_hdr.cells[0]
        c_full.width = col_w_label
        # Merge across 2 columns
        c_full._tc.get_or_add_tcPr().append(parse_xml(f'<w:gridSpan {nsdecls("w")} w:val="2"/>'))
        c_full._tc.get_or_add_tcPr().append(parse_xml(f'<w:shd {nsdecls("w")} w:fill="0F172A"/>'))
        p_mf_hdr = c_full.paragraphs[0]
        r_mf_hdr = p_mf_hdr.add_run("LUỒNG SỰ KIỆN CHÍNH (MAIN SUCCESS FLOW)")
        r_mf_hdr.bold = True
        r_mf_hdr.font.size = Pt(10)
        r_mf_hdr.font.color.rgb = RGBColor(255, 255, 255)

        for step in uc.main_flow:
            row_st = table.add_row()
            row_st._tr.get_or_add_trPr().append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))
            cell_step, cell_act = row_st.cells[0], row_st.cells[1]
            cell_step.width, cell_act.width = col_w_label, col_w_val

            for cell, bg in [(cell_step, "F8FAFC"), (cell_act, "FFFFFF")]:
                c_pr = cell._tc.get_or_add_tcPr()
                c_pr.append(parse_xml(f'<w:shd {nsdecls("w")} w:fill="{bg}"/>'))
                c_pr.append(parse_xml(f'<w:vAlign {nsdecls("w")} w:val="center"/>'))
                c_pr.append(parse_xml(f'<w:tcBorders {nsdecls("w")}><w:top w:val="single" w:sz="4" w:color="CBD5E1"/><w:bottom w:val="single" w:sz="4" w:color="CBD5E1"/><w:left w:val="single" w:sz="4" w:color="CBD5E1"/><w:right w:val="single" w:sz="4" w:color="CBD5E1"/></w:tcBorders>'))

            p_st = cell_step.paragraphs[0]
            r_st_num = p_st.add_run(f"Bước {step.step}\n")
            r_st_num.bold = True
            r_st_actor = p_st.add_run(f"[{step.actor}]")
            r_st_actor.font.size = Pt(8.5)
            r_st_actor.font.color.rgb = RGBColor(2, 132, 199) if "hệ thống" not in step.actor.lower() else RGBColor(14, 116, 144)

            p_act = cell_act.paragraphs[0]
            r_act = p_act.add_run(step.action)
            r_act.font.size = Pt(9.5)

        # Alternative flows
        if uc.alternative_flows:
            row_alt_hdr = table.add_row()
            row_alt_hdr._tr.get_or_add_trPr().append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))
            c_alt = row_alt_hdr.cells[0]
            c_alt._tc.get_or_add_tcPr().append(parse_xml(f'<w:gridSpan {nsdecls("w")} w:val="2"/>'))
            c_alt._tc.get_or_add_tcPr().append(parse_xml(f'<w:shd {nsdecls("w")} w:fill="334155"/>'))
            p_alt = c_alt.paragraphs[0]
            r_alt = p_alt.add_run("LUỒNG RẼ NHÁNH / THAY THẾ (ALTERNATIVE FLOWS)")
            r_alt.bold = True
            r_alt.font.size = Pt(9.5)
            r_alt.font.color.rgb = RGBColor(255, 255, 255)

            for alt in uc.alternative_flows:
                row_item = table.add_row()
                c_lbl, c_val = row_item.cells[0], row_item.cells[1]
                c_lbl.width, c_val.width = col_w_label, col_w_val
                for cell in (c_lbl, c_val):
                    c_pr = cell._tc.get_or_add_tcPr()
                    c_pr.append(parse_xml(f'<w:shd {nsdecls("w")} w:fill="F8FAFC"/>'))
                    c_pr.append(parse_xml(f'<w:tcBorders {nsdecls("w")}><w:top w:val="single" w:sz="4" w:color="CBD5E1"/><w:bottom w:val="single" w:sz="4" w:color="CBD5E1"/><w:left w:val="single" w:sz="4" w:color="CBD5E1"/><w:right w:val="single" w:sz="4" w:color="CBD5E1"/></w:tcBorders>'))
                c_lbl.paragraphs[0].add_run(alt.id).bold = True
                p_v = c_val.paragraphs[0]
                p_v.add_run(f"Điều kiện: {alt.condition}\n").bold = True
                for s in alt.steps:
                    p_v.add_run(f"  • {s}\n")

        # Exception flows
        if uc.exception_flows:
            row_exc_hdr = table.add_row()
            row_exc_hdr._tr.get_or_add_trPr().append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))
            c_exc = row_exc_hdr.cells[0]
            c_exc._tc.get_or_add_tcPr().append(parse_xml(f'<w:gridSpan {nsdecls("w")} w:val="2"/>'))
            c_exc._tc.get_or_add_tcPr().append(parse_xml(f'<w:shd {nsdecls("w")} w:fill="881337"/>'))  # Rose dark
            p_exc = c_exc.paragraphs[0]
            r_exc = p_exc.add_run("LUỒNG NGOẠI LỆ & XỬ LÝ LỖI (EXCEPTION FLOWS)")
            r_exc.bold = True
            r_exc.font.size = Pt(9.5)
            r_exc.font.color.rgb = RGBColor(255, 255, 255)

            for exc in uc.exception_flows:
                row_item = table.add_row()
                c_lbl, c_val = row_item.cells[0], row_item.cells[1]
                c_lbl.width, c_val.width = col_w_label, col_w_val
                for cell in (c_lbl, c_val):
                    c_pr = cell._tc.get_or_add_tcPr()
                    c_pr.append(parse_xml(f'<w:shd {nsdecls("w")} w:fill="FFF1F2"/>'))
                    c_pr.append(parse_xml(f'<w:tcBorders {nsdecls("w")}><w:top w:val="single" w:sz="4" w:color="FECDD3"/><w:bottom w:val="single" w:sz="4" w:color="FECDD3"/><w:left w:val="single" w:sz="4" w:color="FECDD3"/><w:right w:val="single" w:sz="4" w:color="FECDD3"/></w:tcBorders>'))
                c_lbl.paragraphs[0].add_run(exc.id).bold = True
                p_v = c_val.paragraphs[0]
                p_v.add_run(f"Sự cố / Lỗi: {exc.condition}\n").bold = True
                for s in exc.steps:
                    p_v.add_run(f"  ⚠ {s}\n")

        # Business Rules
        if uc.business_rules:
            add_spec_row("Quy tắc nghiệp vụ", uc.business_rules)

        os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
        doc.save(output_path)
        return output_path
