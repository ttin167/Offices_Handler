"""
core.ba_hub.story_builder — Production User Story & Gherkin Acceptance Criteria Engine.

Features:
1. Strict BABOK & Gherkin (Given-When-Then) parsing & validation.
2. BA Quality Audit Linter (evaluates Role specificity, Business Value clarity, and Scenario coverage).
3. Dual-Format Exporter: Clean GitHub Markdown & Publication-Grade OpenXML Word (.docx).
"""

from __future__ import annotations

import os
import re
from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional

import docx
from docx import Document
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.shared import Inches, Pt, RGBColor


@dataclass
class GherkinScenario:
    title: str
    scenario_type: str = "Happy Path"  # Happy Path, Negative Path, Boundary, Edge Case
    given: str = ""
    when: str = ""
    then: str = ""
    and_conditions: List[str] = field(default_factory=list)

    def to_dict(self) -> Dict[str, Any]:
        return {
            "title": self.title,
            "type": self.scenario_type,
            "given": self.given,
            "when": self.when,
            "then": self.then,
            "and": " ".join(self.and_conditions) if isinstance(self.and_conditions, list) else str(self.and_conditions),
        }

    def to_gherkin_text(self) -> str:
        lines = [f"  Scenario: {self.title} [{self.scenario_type}]"]
        if self.given:
            lines.append(f"    Given {self.given}")
        if self.when:
            lines.append(f"    When {self.when}")
        if self.then:
            lines.append(f"    Then {self.then}")
        for cond in self.and_conditions:
            lines.append(f"    And {cond}")
        return "\n".join(lines)


@dataclass
class UserStory:
    id: str
    title: str
    role: str
    goal: str
    benefit: str
    domain: str = "General Business"
    business_rules: List[str] = field(default_factory=list)
    scenarios: List[GherkinScenario] = field(default_factory=list)

    def to_dict(self) -> Dict[str, Any]:
        return {
            "id": self.id,
            "title": self.title,
            "domain": self.domain,
            "role": self.role,
            "goal": self.goal,
            "benefit": self.benefit,
            "business_rules": self.business_rules,
            "scenarios": [s.to_dict() for s in self.scenarios],
        }

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> UserStory:
        return StoryBuilder.parse_story_dict(data)

    def format_as_gherkin(self) -> str:
        out = [
            f"Feature: {self.id} - {self.title}",
            f"  As a {self.role}",
            f"  I want {self.goal}",
            f"  So that {self.benefit}",
            ""
        ]
        if self.business_rules:
            out.append("  # Business Rules:")
            for br in self.business_rules:
                out.append(f"  # - {br}")
            out.append("")

        for sc in self.scenarios:
            out.append(sc.to_gherkin_text())
            out.append("")
        return "\n".join(out)


class StoryQualityLinter:
    """Automated Business Analyst Linting & Quality Assessment for Requirements."""

    GENERIC_ROLES = {"user", "người dùng", "ai đó", "someone", "admin", "client", "customer"}

    @classmethod
    def audit_story(cls, story: Any) -> Dict[str, Any]:
        if hasattr(story, "to_dict"):
            story = story.to_dict()
        elif not isinstance(story, dict):
            story = {}
        score = 100
        findings: List[Dict[str, str]] = []

        role = str(story.get("role", "")).strip().lower()
        goal = str(story.get("goal", "")).strip()
        benefit = str(story.get("benefit", "")).strip()
        scenarios = story.get("scenarios", [])

        # 1. Audit Role Specificity
        if not role:
            score -= 25
            findings.append({"severity": "CRITICAL", "message": "Thiếu thông tin vai trò người dùng (Role) trong cấu trúc 'As a [Role]'."})
        elif role in cls.GENERIC_ROLES:
            score -= 10
            findings.append({
                "severity": "WARNING",
                "message": f"Vai trò '{role}' quá chung chung. Khuyến nghị định danh rõ Persona nghiệp vụ (ví dụ: 'Khách hàng cá nhân đã xác thực eKYC', 'Nhân viên kế toán công nợ')."
            })

        # 2. Audit Goal & Solution Bias
        if not goal:
            score -= 25
            findings.append({"severity": "CRITICAL", "message": "Thiếu mục tiêu tính năng (Goal) trong cấu trúc 'I want [Goal]'."})
        elif any(tech in goal.lower() for tech in ["bấm vào nút", "click button", "sql", "api", "database", "mã nguồn"]):
            score -= 10
            findings.append({
                "severity": "INFO",
                "message": "Mục tiêu (Goal) có dấu hiệu ép giải pháp kỹ thuật hoặc chi tiết giao diện (Solution Bias). Nên tập trung vào nhu cầu nghiệp vụ."
            })

        # 3. Audit Business Value / Benefit
        if not benefit:
            score -= 20
            findings.append({"severity": "CRITICAL", "message": "Thiếu giá trị nghiệp vụ (Benefit) trong cấu trúc 'So that [Benefit]'."})

        # 4. Audit Acceptance Criteria & Scenarios
        if not scenarios:
            score -= 30
            findings.append({"severity": "CRITICAL", "message": "Chưa có tiêu chí nghiệm thu Acceptance Criteria (Gherkin Scenarios) nào được định nghĩa."})
        else:
            has_happy = any(s.get("type", "").lower() in ["happy path", "happy"] or "thành công" in s.get("title", "").lower() for s in scenarios)
            has_negative = any(s.get("type", "").lower() in ["negative path", "negative", "exception", "boundary"] or "lỗi" in s.get("title", "").lower() or "từ chối" in s.get("title", "").lower() for s in scenarios)

            if not has_happy:
                score -= 10
                findings.append({"severity": "WARNING", "message": "Thiếu kịch bản luồng thành công chính (Happy Path Scenario)."})
            if not has_negative:
                score -= 15
                findings.append({"severity": "WARNING", "message": "Thiếu kịch bản xử lý lỗi hoặc ngoại lệ (Negative Path / Boundary Scenario). Hệ thống có nguy cơ bỏ sót validation."})

            # Check Given-When-Then completeness
            for idx, sc in enumerate(scenarios, start=1):
                missing_kw = []
                if not sc.get("given"): missing_kw.append("Given")
                if not sc.get("when"): missing_kw.append("When")
                if not sc.get("then"): missing_kw.append("Then")
                if missing_kw:
                    score -= 5
                    findings.append({
                        "severity": "WARNING",
                        "message": f"Kịch bản #{idx} ('{sc.get('title', 'Không tên')}') thiếu từ khóa chuẩn Gherkin: {', '.join(missing_kw)}."
                    })

        final_score = max(0, min(100, score))
        grade = "Xuất sắc (A+)" if final_score >= 90 else "Tốt (B)" if final_score >= 75 else "Cần hoàn thiện (C)" if final_score >= 50 else "Không đạt (D)"

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
            "total_scenarios": len(scenarios)
        })

    @classmethod
    def lint_story(cls, story: Any) -> Any:
        return cls.audit_story(story)


class StoryBuilder:
    """Builds, validates, and exports Business Analyst User Stories."""

    @staticmethod
    def parse_story_dict(data: Dict[str, Any]) -> UserStory:
        story_id = data.get("id") or "US-001"
        title = data.get("title") or "Yêu cầu nghiệp vụ mới"
        domain = data.get("domain") or "General Business"
        role = data.get("role") or "Người dùng"
        goal = data.get("goal") or "thao tác chức năng"
        benefit = data.get("benefit") or "đáp ứng nhu cầu công việc"
        rules = data.get("business_rules", [])

        scenarios: List[GherkinScenario] = []
        for raw_sc in data.get("scenarios", []):
            and_conds = raw_sc.get("and", [])
            if isinstance(and_conds, str):
                and_conds = [and_conds] if and_conds.strip() else []
            sc = GherkinScenario(
                title=raw_sc.get("title", "Kịch bản kiểm thử"),
                scenario_type=raw_sc.get("type", "Happy Path"),
                given=raw_sc.get("given", ""),
                when=raw_sc.get("when", ""),
                then=raw_sc.get("then", ""),
                and_conditions=and_conds,
            )
            scenarios.append(sc)

        return UserStory(
            id=story_id,
            title=title,
            domain=domain,
            role=role,
            goal=goal,
            benefit=benefit,
            business_rules=rules,
            scenarios=scenarios,
        )

    @classmethod
    def export_to_markdown(cls, stories: List[UserStory], output_path: Optional[str] = None) -> str:
        doc = ["# TÀI LIỆU ĐẶC TẢ USER STORIES & ACCEPTANCE CRITERIA (GHERKIN)", ""]
        for st in stories:
            doc.append(f"## {st.id}: {st.title}")
            doc.append(f"**Phân hệ / Domain:** `{st.domain}`")
            doc.append("")
            doc.append("> **As a** " + st.role + "  ")
            doc.append("> **I want** " + st.goal + "  ")
            doc.append("> **So that** " + st.benefit)
            doc.append("")
            if st.business_rules:
                doc.append("### Quy tắc nghiệp vụ (Business Rules):")
                for br in st.business_rules:
                    doc.append(f"- {br}")
                doc.append("")

            doc.append("### Tiêu chí nghiệm thu (Acceptance Criteria - Gherkin Scenarios):")
            for idx, sc in enumerate(st.scenarios, start=1):
                doc.append(f"#### Kịch bản {idx}: {sc.title} *({sc.scenario_type})*")
                doc.append("```gherkin")
                doc.append(sc.to_gherkin_text())
                doc.append("```")
                doc.append("")
            doc.append("---")
            doc.append("")
        md_text = "\n".join(doc)
        if output_path:
            with open(output_path, "w", encoding="utf-8") as f:
                f.write(md_text)
        return md_text

    @classmethod
    def export_to_docx(cls, stories: List[UserStory], output_path: str, doc_title: Optional[str] = None) -> str:
        doc = Document()

        # Page setup A4
        for sec in doc.sections:
            sec.top_margin = Inches(0.8)
            sec.bottom_margin = Inches(0.8)
            sec.left_margin = Inches(1.0)
            sec.right_margin = Inches(0.8)

        # Title
        p_title = doc.add_paragraph()
        p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        title_text = doc_title or "TÀI LIỆU ĐẶC TẢ USER STORIES & ACCEPTANCE CRITERIA"
        r_title = p_title.add_run(title_text)
        r_title.bold = True
        r_title.font.size = Pt(18)
        r_title.font.color.rgb = RGBColor(15, 23, 42)  # Dark slate

        p_sub = doc.add_paragraph()
        p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r_sub = p_sub.add_run("Chuẩn Đặc Tả Nghiệp Vụ Cấp Doanh Nghiệp (BABOK / Agile Gherkin Spec)")
        r_sub.italic = True
        r_sub.font.size = Pt(10.5)
        r_sub.font.color.rgb = RGBColor(100, 116, 139)

        doc.add_paragraph().paragraph_format.space_after = Pt(12)

        for story in stories:
            # Story Header
            h2 = doc.add_heading(level=1)
            h2_run = h2.add_run(f"{story.id} — {story.title}")
            h2_run.font.size = Pt(14)
            h2_run.font.color.rgb = RGBColor(30, 41, 59)

            p_dom = doc.add_paragraph()
            p_dom.paragraph_format.space_after = Pt(6)
            r_dom_label = p_dom.add_run("Phân hệ nghiệp vụ: ")
            r_dom_label.bold = True
            r_dom_val = p_dom.add_run(story.domain)
            r_dom_val.font.color.rgb = RGBColor(14, 116, 144)

            # User Story Statement Box (Table with slate border & background)
            t_story = doc.add_table(rows=1, cols=1)
            t_story.alignment = WD_TABLE_ALIGNMENT.CENTER
            t_story.autofit = False
            cell = t_story.cell(0, 0)
            cell.width = Inches(6.5)

            # OpenXML shading & padding
            tcPr = cell._tc.get_or_add_tcPr()
            shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="F1F5F9"/>')
            tcBorders = parse_xml(f'<w:tcBorders {nsdecls("w")}><w:left w:val="single" w:sz="24" w:space="0" w:color="0284C7"/><w:top w:val="none"/><w:right w:val="none"/><w:bottom w:val="none"/></w:tcBorders>')
            tcPr.append(shd)
            tcPr.append(tcBorders)

            p_box = cell.paragraphs[0]
            p_box.paragraph_format.space_before = Pt(6)
            p_box.paragraph_format.space_after = Pt(6)

            r_as = p_box.add_run("As a ")
            r_as.bold = True
            r_as.font.color.rgb = RGBColor(15, 23, 42)
            p_box.add_run(f"{story.role}\n")

            r_want = p_box.add_run("I want ")
            r_want.bold = True
            r_want.font.color.rgb = RGBColor(15, 23, 42)
            p_box.add_run(f"{story.goal}\n")

            r_so = p_box.add_run("So that ")
            r_so.bold = True
            r_so.font.color.rgb = RGBColor(15, 23, 42)
            p_box.add_run(story.benefit)

            doc.add_paragraph().paragraph_format.space_after = Pt(8)

            # Business rules
            if story.business_rules:
                p_br = doc.add_paragraph()
                p_br.add_run("Ràng Buộc Nghiệp Vụ (Business Rules):").bold = True
                for br in story.business_rules:
                    p_item = doc.add_paragraph(style='List Bullet')
                    p_item.add_run(br)

            # Acceptance Criteria Scenarios Table
            if story.scenarios:
                p_ac = doc.add_paragraph()
                p_ac.add_run("Tiêu Chí Nghiệm Thu (Acceptance Criteria - Gherkin Matrix):").bold = True

                t_sc = doc.add_table(rows=1, cols=4)
                t_sc.alignment = WD_TABLE_ALIGNMENT.CENTER
                hdr = t_sc.rows[0]

                headers = ["STT / Loại", "Kịch Bản (Scenario)", "Điều Kiện (Given / When)", "Kết Quả Mong Đợi (Then / And)"]
                col_widths = [Inches(1.1), Inches(1.8), Inches(1.8), Inches(1.8)]

                # Invariant: tblHeader & cantSplit
                trPr = hdr._tr.get_or_add_trPr()
                trPr.append(parse_xml(f'<w:tblHeader {nsdecls("w")}/>'))
                trPr.append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))

                for idx, text in enumerate(headers):
                    c = hdr.cells[idx]
                    c.width = col_widths[idx]
                    c_pr = c._tc.get_or_add_tcPr()
                    c_pr.append(parse_xml(f'<w:shd {nsdecls("w")} w:fill="1E293B"/>'))
                    c_pr.append(parse_xml(f'<w:vAlign {nsdecls("w")} w:val="center"/>'))
                    p = c.paragraphs[0]
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    r = p.add_run(text)
                    r.bold = True
                    r.font.size = Pt(9.5)
                    r.font.color.rgb = RGBColor(255, 255, 255)

                for idx, sc in enumerate(story.scenarios, start=1):
                    row = t_sc.add_row()
                    row_tr = row._tr.get_or_add_trPr()
                    row_tr.append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))

                    # Cell 0: Index & Type
                    c0 = row.cells[0]
                    c0.width = col_widths[0]
                    p0 = c0.paragraphs[0]
                    p0.add_run(f"#{idx}\n").bold = True
                    r_type = p0.add_run(sc.scenario_type)
                    r_type.font.size = Pt(8.5)
                    r_type.font.color.rgb = RGBColor(2, 132, 199) if "happy" in sc.scenario_type.lower() else RGBColor(225, 29, 72)

                    # Cell 1: Scenario Title
                    c1 = row.cells[1]
                    c1.width = col_widths[1]
                    p1 = c1.paragraphs[0]
                    p1.add_run(sc.title).bold = True
                    p1.runs[0].font.size = Pt(9)

                    # Cell 2: Given / When
                    c2 = row.cells[2]
                    c2.width = col_widths[2]
                    p2 = c2.paragraphs[0]
                    p2.add_run("Given: ").bold = True
                    p2.add_run(f"{sc.given}\n\n")
                    p2.add_run("When: ").bold = True
                    p2.add_run(sc.when)
                    for r in p2.runs: r.font.size = Pt(8.5)

                    # Cell 3: Then / And
                    c3 = row.cells[3]
                    c3.width = col_widths[3]
                    p3 = c3.paragraphs[0]
                    p3.add_run("Then: ").bold = True
                    p3.add_run(f"{sc.then}")
                    if sc.and_conditions:
                        p3.add_run("\n\nAnd: ").bold = True
                        p3.add_run("; ".join(sc.and_conditions))
                    for r in p3.runs: r.font.size = Pt(8.5)

                    # Zebra striping
                    if idx % 2 == 0:
                        for cell_item in [c0, c1, c2, c3]:
                            cell_pr = cell_item._tc.get_or_add_tcPr()
                            cell_pr.append(parse_xml(f'<w:shd {nsdecls("w")} w:fill="F8FAFC"/>'))

            doc.add_paragraph().paragraph_format.space_after = Pt(16)

        os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
        doc.save(output_path)
        return output_path
