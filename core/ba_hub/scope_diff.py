"""
core.ba_hub.scope_diff — Scope Creep & Change Impact Analysis Engine.

Built on top of OpenXML Legal Redline Diff (core/docx/diff_engine.py):
1. Semantic Delta Parser:
   Categorizes changes into NEW_REQUIREMENT, MODIFIED_LOGIC, REMOVED_SCOPE, DATA_FIELD_DELTA.
2. Impact Severity Matrix & Effort Estimator:
   Computes estimated Man-days / Effort for Change Requests (CR).
3. Publication-Grade Word (.docx) Scope Creep Impact Report.
"""

from __future__ import annotations

import os
import re
import xml.etree.ElementTree as ET
import zipfile
from dataclasses import dataclass, field
from datetime import datetime
from typing import Any, Dict, List, Optional, Tuple

import docx
from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
from docx.shared import Inches, Pt, RGBColor

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


@dataclass
class ScopeDeltaItem:
    category: str  # NEW_REQUIREMENT, MODIFIED_LOGIC, REMOVED_SCOPE, DATA_FIELD_DELTA
    severity: str  # CRITICAL, MAJOR, MINOR, TRIVIAL
    section_title: str
    original_text: str
    revised_text: str
    impact_description: str
    estimated_mandays: float = 0.0

    def to_dict(self) -> Dict[str, Any]:
        return {
            "category": self.category,
            "severity": self.severity,
            "section": self.section_title,
            "original": self.original_text,
            "revised": self.revised_text,
            "impact": self.impact_description,
            "estimated_mandays": self.estimated_mandays,
        }


@dataclass
class ScopeChangeSummary:
    total_changes: int
    new_requirements_count: int
    modified_logic_count: int
    removed_scope_count: int
    data_field_changes_count: int
    total_estimated_mandays: float
    impact_level: str  # HIGH, MEDIUM, LOW
    deltas: List[ScopeDeltaItem] = field(default_factory=list)

    def to_dict(self) -> Dict[str, Any]:
        return {
            "total_changes": self.total_changes,
            "new_requirements_count": self.new_requirements_count,
            "modified_logic_count": self.modified_logic_count,
            "removed_scope_count": self.removed_scope_count,
            "data_field_changes_count": self.data_field_changes_count,
            "total_estimated_mandays": round(self.total_estimated_mandays, 1),
            "impact_level": self.impact_level,
            "deltas": [d.to_dict() for d in self.deltas],
        }


class ScopeCreepAuditor:
    """Analyzes differences between two Word SRS/BRD versions and computes scope impact."""

    @classmethod
    def analyze_documents(
        cls,
        base_doc_path: str,
        revised_doc_path: str,
    ) -> ScopeChangeSummary:
        # Read text & structure from both documents
        base_paras = cls._extract_paragraphs(base_doc_path)
        rev_paras = cls._extract_paragraphs(revised_doc_path)

        deltas: List[ScopeDeltaItem] = []
        base_set = {p["text"].strip(): p for p in base_paras if p["text"].strip()}
        rev_set = {p["text"].strip(): p for p in rev_paras if p["text"].strip()}

        # 1. Detect New Requirements (in revised but not in base)
        for r_text, r_info in rev_set.items():
            if r_text not in base_set:
                # Check if it's a modification of a similar sentence
                is_mod = False
                for b_text in base_set:
                    sim = cls._similarity(r_text, b_text)
                    if 0.4 < sim < 0.95:
                        is_mod = True
                        cat = "MODIFIED_LOGIC"
                        sev, mandays = cls._evaluate_severity(r_text, b_text)
                        deltas.append(
                            ScopeDeltaItem(
                                category=cat,
                                severity=sev,
                                section_title=r_info["heading"] or "Quy tắc nghiệp vụ",
                                original_text=b_text,
                                revised_text=r_text,
                                impact_description=f"Sửa đổi logic: {sev} impact",
                                estimated_mandays=mandays,
                            )
                        )
                        break

                if not is_mod and len(r_text) > 20:
                    cat = "NEW_REQUIREMENT"
                    sev, mandays = cls._evaluate_new_req_severity(r_text)
                    deltas.append(
                        ScopeDeltaItem(
                            category=cat,
                            severity=sev,
                            section_title=r_info["heading"] or "Phạm vi mới",
                            original_text="-",
                            revised_text=r_text,
                            impact_description=f"Yêu cầu bổ sung ngoài phạm vi ban đầu: {sev} impact",
                            estimated_mandays=mandays,
                        )
                    )

        # 2. Detect Removed Scope (in base but deleted from revised)
        for b_text, b_info in base_set.items():
            if b_text not in rev_set:
                already_matched = any(d.original_text == b_text for d in deltas)
                if not already_matched and len(b_text) > 20:
                    deltas.append(
                        ScopeDeltaItem(
                            category="REMOVED_SCOPE",
                            severity="MINOR",
                            section_title=b_info["heading"] or "Phạm vi cắt giảm",
                            original_text=b_text,
                            revised_text="-",
                            impact_description="Yêu cầu bị loại bỏ khỏi phạm vi",
                            estimated_mandays=0.0,
                        )
                    )

        # Summarize
        new_cnt = sum(1 for d in deltas if d.category == "NEW_REQUIREMENT")
        mod_cnt = sum(1 for d in deltas if d.category == "MODIFIED_LOGIC")
        rem_cnt = sum(1 for d in deltas if d.category == "REMOVED_SCOPE")
        data_cnt = sum(1 for d in deltas if d.category == "DATA_FIELD_DELTA")
        total_md = sum(d.estimated_mandays for d in deltas)

        impact_level = "HIGH" if total_md >= 10.0 or new_cnt >= 5 else "MEDIUM" if total_md >= 3.0 else "LOW"

        return ScopeChangeSummary(
            total_changes=len(deltas),
            new_requirements_count=new_cnt,
            modified_logic_count=mod_cnt,
            removed_scope_count=rem_cnt,
            data_field_changes_count=data_cnt,
            total_estimated_mandays=total_md,
            impact_level=impact_level,
            deltas=deltas,
        )

    @classmethod
    def _extract_paragraphs(cls, doc_path: str) -> List[Dict[str, str]]:
        if not os.path.exists(doc_path):
            return []
        doc = Document(doc_path)
        result = []
        cur_heading = "Mở đầu"
        for p in doc.paragraphs:
            txt = p.text.strip()
            if p.style and p.style.name and p.style.name.startswith("Heading"):
                cur_heading = txt or cur_heading
            elif txt:
                result.append({"text": txt, "heading": cur_heading})
        return result

    @classmethod
    def _similarity(cls, s1: str, s2: str) -> float:
        w1 = set(s1.lower().split())
        w2 = set(s2.lower().split())
        if not w1 or not w2:
            return 0.0
        return len(w1 & w2) / len(w1 | w2)

    @classmethod
    def _evaluate_severity(cls, rev: str, orig: str) -> Tuple[str, float]:
        rev_l = rev.lower()
        if any(w in rev_l for w in ["thanh toán", "cổng", "bảo mật", "otp", "sinh trắc", "api", "tích hợp"]):
            return ("CRITICAL", 4.0)
        if any(w in rev_l for w in ["hạn mức", "phí", "quy tắc", "thời gian", "ngoại lệ", "bắt buộc"]):
            return ("MAJOR", 2.0)
        if any(w in rev_l for w in ["thông báo", "màu", "nút", "hiển thị"]):
            return ("MINOR", 0.5)
        return ("TRIVIAL", 0.0)

    @classmethod
    def _evaluate_new_req_severity(cls, text: str) -> Tuple[str, float]:
        t_l = text.lower()
        if any(w in t_l for w in ["module", "phân hệ", "tích hợp cổng", "core banking", "hóa đơn điện tử"]):
            return ("CRITICAL", 5.0)
        if any(w in t_l for w in ["kịch bản", "cho phép", "kiểm tra", "báo cáo", "tra cứu"]):
            return ("MAJOR", 2.0)
        return ("MINOR", 0.5)

    @classmethod
    def export_scope_report_docx(
        cls,
        summary: ScopeChangeSummary,
        output_path: str,
        project_name: str = "Hệ Thống Doanh Nghiệp",
    ) -> str:
        doc = Document()
        for sec in doc.sections:
            sec.top_margin = Inches(0.8)
            sec.bottom_margin = Inches(0.8)
            sec.left_margin = Inches(1.0)
            sec.right_margin = Inches(0.8)

        # Title
        p_title = doc.add_paragraph()
        p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p_title.add_run("BÁO CÁO THẨM ĐỊNH THAY ĐỔI PHẠM VI (SCOPE CREEP AUDIT)")
        r.bold = True
        r.font.size = Pt(16)
        r.font.color.rgb = RGBColor(15, 23, 42)

        p_sub = doc.add_paragraph()
        p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_sub.add_run(f"Dự án: {project_name} | Đánh giá mức độ tác động Change Request (CR)").italic = True

        # Executive Metrics Box
        p_kpi = doc.add_paragraph()
        p_kpi.paragraph_format.space_before = Pt(12)
        p_kpi.paragraph_format.space_after = Pt(12)
        r_kpi = p_kpi.add_run(
            f"MỨC ĐỘ TÁC ĐỘNG: {summary.impact_level}  |  "
            f"TỔNG THAY ĐỔI: {summary.total_changes}  |  "
            f"YÊU CẦU MỚI: {summary.new_requirements_count}  |  "
            f"ƯỚC TÍNH MAN-DAYS: {summary.total_estimated_mandays:.1f} MD"
        )
        r_kpi.bold = True
        r_kpi.font.size = Pt(10)
        r_kpi.font.color.rgb = RGBColor(185, 28, 28) if summary.impact_level == "HIGH" else RGBColor(15, 23, 42)

        # Delta Table
        table = doc.add_table(rows=1, cols=5)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        hdr = table.rows[0]
        hdr._tr.get_or_add_trPr().append(parse_xml(f'<w:tblHeader {nsdecls("w")}/>'))

        headers = ["Phân Loại", "Mức Độ", "Mục Tiêu Nghiệp Vụ", "Nội Dung Thay Đổi", "Ước Tính (MD)"]
        widths = [Inches(1.5), Inches(1.0), Inches(1.6), Inches(2.3), Inches(0.8)]

        for idx, text in enumerate(headers):
            c = hdr.cells[idx]
            c.width = widths[idx]
            c._tc.get_or_add_tcPr().append(parse_xml(f'<w:shd {nsdecls("w")} w:fill="0F172A"/>'))
            p = c.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = p.add_run(text)
            run.bold = True
            run.font.size = Pt(9)
            run.font.color.rgb = RGBColor(255, 255, 255)

        for d in summary.deltas:
            row = table.add_row()
            row._tr.get_or_add_trPr().append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))
            diff_text = d.revised_text if d.category == "NEW_REQUIREMENT" else f"- Cũ: {d.original_text[:80]}...\n+ Mới: {d.revised_text[:80]}..."
            vals = [
                d.category,
                d.severity,
                d.section_title[:30],
                diff_text,
                f"{d.estimated_mandays:.1f}",
            ]

            for c_idx, val in enumerate(vals):
                cell = row.cells[c_idx]
                cell.width = widths[c_idx]
                cell._tc.get_or_add_tcPr().append(
                    parse_xml(
                        f'<w:tcBorders {nsdecls("w")}><w:top w:val="single" w:sz="4" w:color="CBD5E1"/><w:bottom w:val="single" w:sz="4" w:color="CBD5E1"/><w:left w:val="single" w:sz="4" w:color="CBD5E1"/><w:right w:val="single" w:sz="4" w:color="CBD5E1"/></w:tcBorders>'
                    )
                )
                cell.paragraphs[0].add_run(val).font.size = Pt(8.5)

        os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
        doc.save(output_path)
        return output_path
