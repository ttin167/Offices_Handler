"""
core.ba_hub.decision_table — BABOK v3 Decision Table & Matrix Generator for Business Analysts.

Features:
1. Declarative Decision Table Specification (Conditions, Actions, and Rules).
2. Automated Completeness & Contradiction Audit:
   - Completeness Check: Computes Cartesian product of conditions, identifies missing edge-case rules.
   - Contradiction Check: Detects identical condition states leading to conflicting outcomes.
   - Redundancy Check: Detects duplicate rules.
3. High-Fidelity OpenXML Word (.docx) & Excel (.xlsx) export with decoupled color coding:
   - Condition rows: Light Cyan/Slate (#E0F2FE / #F0F9FF)
   - Action rows: Light Emerald (#DCFCE7 / #F0FDF4)
   - Table invariants: <w:cantSplit/>, <w:tblHeader/>, <w:vAlign w:val="center"/>
"""

from __future__ import annotations

import datetime
import itertools
import os
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Dict, List, Optional, Set, Tuple, Union

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

try:
    from docx import Document
    from docx.shared import Inches, Pt, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.enum.table import WD_TABLE_ALIGNMENT
    from docx.oxml import parse_xml
    from docx.oxml.ns import nsdecls
except ImportError:
    Document = None


@dataclass
class ConditionRule:
    name: str
    description: str = ""
    possible_values: List[str] = field(default_factory=lambda: ["Y", "N"])


@dataclass
class ActionRule:
    name: str
    description: str = ""
    possible_values: List[str] = field(default_factory=lambda: ["X", "-"])


@dataclass
class DecisionRuleEntry:
    rule_id: str
    conditions: Dict[str, str]  # condition_name -> value
    actions: Dict[str, str]     # action_name -> value
    notes: str = ""


@dataclass
class DecisionAuditResult:
    is_valid: bool
    score: int
    total_theoretical_rules: int
    actual_rules_count: int
    missing_combinations: List[Dict[str, str]] = field(default_factory=list)
    contradictions: List[str] = field(default_factory=list)
    redundancies: List[str] = field(default_factory=list)

    def to_dict(self) -> Dict[str, Any]:
        return {
            "is_valid": self.is_valid,
            "score": self.score,
            "theoretical_rules": self.total_theoretical_rules,
            "actual_rules": self.actual_rules_count,
            "missing_combinations": self.missing_combinations[:10],
            "contradictions": self.contradictions,
            "redundancies": self.redundancies,
        }


class DecisionTable:
    """Enterprise Decision Table model with BABOK v3 validation and export engines."""

    def __init__(
        self,
        table_id: str = "DT-001",
        title: str = "Bảng Quyết Định Nghiệp Vụ",
        domain: str = "Chung",
        description: str = "",
        conditions: Optional[List[ConditionRule]] = None,
        actions: Optional[List[ActionRule]] = None,
        rules: Optional[List[DecisionRuleEntry]] = None,
    ):
        self.table_id = table_id
        self.title = title
        self.domain = domain
        self.description = description
        self.conditions: List[ConditionRule] = conditions or []
        self.actions: List[ActionRule] = actions or []
        self.rules: List[DecisionRuleEntry] = rules or []

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> "DecisionTable":
        t_id = data.get("id") or data.get("table_id") or "DT-001"
        title = data.get("title") or "Bảng Quyết Định Nghiệp Vụ"
        domain = data.get("domain") or "General Business"
        desc = data.get("description") or ""

        conds = []
        for c in data.get("conditions", []):
            if isinstance(c, str):
                conds.append(ConditionRule(name=c))
            elif isinstance(c, dict):
                conds.append(ConditionRule(
                    name=c.get("name", ""),
                    description=c.get("description", ""),
                    possible_values=c.get("values") or c.get("possible_values") or ["Y", "N"]
                ))

        acts = []
        for a in data.get("actions", []):
            if isinstance(a, str):
                acts.append(ActionRule(name=a))
            elif isinstance(a, dict):
                acts.append(ActionRule(
                    name=a.get("name", ""),
                    description=a.get("description", ""),
                    possible_values=a.get("values") or a.get("possible_values") or ["X", "-"]
                ))

        rules = []
        for idx, r in enumerate(data.get("rules", []), start=1):
            r_id = r.get("rule_id") or f"R{idx:02d}"
            rules.append(DecisionRuleEntry(
                rule_id=r_id,
                conditions=r.get("conditions", {}),
                actions=r.get("actions", {}),
                notes=r.get("notes", "")
            ))

        return cls(table_id=t_id, title=title, domain=domain, description=desc, conditions=conds, actions=acts, rules=rules)

    def to_dict(self) -> Dict[str, Any]:
        return {
            "table_id": self.table_id,
            "title": self.title,
            "domain": self.domain,
            "description": self.description,
            "conditions": [{"name": c.name, "description": c.description, "values": c.possible_values} for c in self.conditions],
            "actions": [{"name": a.name, "description": a.description, "values": a.possible_values} for a in self.actions],
            "rules": [
                {
                    "rule_id": r.rule_id,
                    "conditions": r.conditions,
                    "actions": r.actions,
                    "notes": r.notes
                }
                for r in self.rules
            ]
        }

    def audit_completeness_and_contradictions(self) -> DecisionAuditResult:
        """Audits decision table completeness, contradictions, and redundancies per BABOK v3."""
        if not self.conditions:
            return DecisionAuditResult(is_valid=True, score=100, total_theoretical_rules=0, actual_rules_count=0)

        # 1. Cartesian product of all condition values
        cond_names = [c.name for c in self.conditions]
        cond_value_lists = [c.possible_values for c in self.conditions]
        
        all_combinations = list(itertools.product(*cond_value_lists))
        total_theoretical = len(all_combinations)

        # Build mapping of condition tuples to rules
        seen_conditions: Dict[Tuple[str, ...], List[DecisionRuleEntry]] = {}
        contradictions = []
        redundancies = []

        for rule in self.rules:
            # Condition vector
            cond_tuple = tuple(str(rule.conditions.get(c_name, "-")).strip() for c_name in cond_names)
            if cond_tuple in seen_conditions:
                # Compare actions
                prev_rules = seen_conditions[cond_tuple]
                for prev in prev_rules:
                    if prev.actions == rule.actions:
                        redundancies.append(f"Quy tắc {rule.rule_id} trùng lặp hoàn toàn với {prev.rule_id}.")
                    else:
                        contradictions.append(
                            f"Mâu thuẫn logic giữa {prev.rule_id} và {rule.rule_id}: Cùng điều kiện {dict(zip(cond_names, cond_tuple))} nhưng hành động khác nhau."
                        )
                seen_conditions[cond_tuple].append(rule)
            else:
                seen_conditions[cond_tuple] = [rule]

        # 2. Check for missing combinations
        missing = []
        for comb in all_combinations:
            comb_tuple = tuple(str(v).strip() for v in comb)
            if comb_tuple not in seen_conditions:
                # Check wildcard '-' match
                matched = False
                for existing_tuple in seen_conditions:
                    matches_all = True
                    for ex_val, req_val in zip(existing_tuple, comb_tuple):
                        if ex_val != "-" and ex_val != req_val:
                            matches_all = False
                            break
                    if matches_all:
                        matched = True
                        break
                if not matched:
                    missing.append(dict(zip(cond_names, comb_tuple)))

        score = 100
        if contradictions:
            score -= min(50, len(contradictions) * 20)
        if redundancies:
            score -= min(20, len(redundancies) * 10)
        if total_theoretical > 0 and missing:
            missing_ratio = len(missing) / total_theoretical
            score -= int(min(40, missing_ratio * 40))

        score = max(0, min(100, score))
        is_valid = len(contradictions) == 0 and score >= 70

        return DecisionAuditResult(
            is_valid=is_valid,
            score=score,
            total_theoretical_rules=total_theoretical,
            actual_rules_count=len(self.rules),
            missing_combinations=missing,
            contradictions=contradictions,
            redundancies=redundancies
        )

    def export_to_excel(self, output_path: str) -> str:
        """Exports decision table to styled Excel OpenXML sheet."""
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Decision_Table"
        ws.views.sheetView[0].showGridLines = True

        # Header Title
        num_rules = len(self.rules)
        total_cols = max(4, 2 + num_rules)
        end_col_letter = openpyxl.utils.get_column_letter(total_cols)

        ws.merge_cells(f"A1:{end_col_letter}1")
        title_cell = ws["A1"]
        title_cell.value = f"BẢNG QUYẾT ĐỊNH NGHIỆP VỤ (DECISION TABLE) — {self.table_id}: {self.title.upper()}"
        title_cell.font = Font(name="Segoe UI", size=13, bold=True, color="FFFFFF")
        title_cell.fill = PatternFill(start_color="0F172A", end_color="0F172A", fill_type="solid")
        title_cell.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[1].height = 36

        # Subtitle
        ws.merge_cells(f"A2:{end_col_letter}2")
        sub_cell = ws["A2"]
        sub_cell.value = f"Phân hệ: {self.domain} | Chuẩn BABOK v3 | Xuất bản: {datetime.date.today().strftime('%d/%m/%Y')}"
        sub_cell.font = Font(name="Segoe UI", size=9.5, italic=True, color="64748B")
        sub_cell.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[2].height = 20

        # Table Header Row
        headers = ["Phân Loại", "Điều Kiện / Hành Động Nghiệp Vụ"] + [r.rule_id for r in self.rules]
        ws.row_dimensions[4].height = 28
        h_fill = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")
        h_font = Font(name="Segoe UI", size=10, bold=True, color="FFFFFF")
        for col_idx, h_text in enumerate(headers, start=1):
            cell = ws.cell(row=4, column=col_idx, value=h_text)
            cell.fill = h_fill
            cell.font = h_font
            cell.alignment = Alignment(horizontal="center", vertical="center")

        thin_border = Border(
            top=Side(style="thin", color="CBD5E1"), bottom=Side(style="thin", color="CBD5E1"),
            left=Side(style="thin", color="CBD5E1"), right=Side(style="thin", color="CBD5E1")
        )

        curr_row = 5
        # Conditions Section
        cond_fill = PatternFill(start_color="F0F9FF", end_color="F0F9FF", fill_type="solid")
        for idx, cond in enumerate(self.conditions, start=1):
            ws.row_dimensions[curr_row].height = 24
            c_tag = ws.cell(row=curr_row, column=1, value="ĐIỀU KIỆN" if idx == 1 else "")
            c_tag.font = Font(name="Segoe UI", size=9.5, bold=True, color="0369A1")
            c_tag.alignment = Alignment(horizontal="center", vertical="center")

            c_name = ws.cell(row=curr_row, column=2, value=cond.name)
            c_name.font = Font(name="Segoe UI", size=9.5, bold=True, color="0F172A")
            c_name.alignment = Alignment(horizontal="left", vertical="center")

            for r_idx, rule in enumerate(self.rules, start=1):
                val = str(rule.conditions.get(cond.name, "-"))
                val_cell = ws.cell(row=curr_row, column=2 + r_idx, value=val)
                val_cell.font = Font(name="Consolas", size=10, bold=True, color="0284C7" if val in ("Y", "YES", "TRUE", "X") else "64748B")
                val_cell.fill = cond_fill
                val_cell.alignment = Alignment(horizontal="center", vertical="center")

            for col in range(1, total_cols + 1):
                ws.cell(row=curr_row, column=col).border = thin_border
            curr_row += 1

        # Divider Separator
        ws.row_dimensions[curr_row].height = 6
        for col in range(1, total_cols + 1):
            sep = ws.cell(row=curr_row, column=col)
            sep.fill = PatternFill(start_color="0284C7", end_color="0284C7", fill_type="solid")
        curr_row += 1

        # Actions Section
        act_fill = PatternFill(start_color="F0FDF4", end_color="F0FDF4", fill_type="solid")
        for idx, act in enumerate(self.actions, start=1):
            ws.row_dimensions[curr_row].height = 24
            a_tag = ws.cell(row=curr_row, column=1, value="HÀNH ĐỘNG" if idx == 1 else "")
            a_tag.font = Font(name="Segoe UI", size=9.5, bold=True, color="16A34A")
            a_tag.alignment = Alignment(horizontal="center", vertical="center")

            a_name = ws.cell(row=curr_row, column=2, value=act.name)
            a_name.font = Font(name="Segoe UI", size=9.5, bold=True, color="0F172A")
            a_name.alignment = Alignment(horizontal="left", vertical="center")

            for r_idx, rule in enumerate(self.rules, start=1):
                val = str(rule.actions.get(act.name, "-"))
                val_cell = ws.cell(row=curr_row, column=2 + r_idx, value=val)
                val_cell.font = Font(name="Consolas", size=10, bold=True, color="15803D" if val in ("X", "YES", "EXECUTE", "DUYET") else "64748B")
                val_cell.fill = act_fill
                val_cell.alignment = Alignment(horizontal="center", vertical="center")

            for col in range(1, total_cols + 1):
                ws.cell(row=curr_row, column=col).border = thin_border
            curr_row += 1

        # Set column widths
        ws.column_dimensions["A"].width = 14
        ws.column_dimensions["B"].width = 36
        for r_idx in range(1, num_rules + 1):
            col_letter = openpyxl.utils.get_column_letter(2 + r_idx)
            ws.column_dimensions[col_letter].width = 10

        ws.freeze_panes = "C5"
        os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
        wb.save(output_path)
        return output_path

    def export_to_docx(self, output_path: str) -> str:
        """Exports decision table to OpenXML Word document with table invariants."""
        if Document is None:
            raise RuntimeError("Thư viện python-docx chưa được cài đặt.")

        doc = Document()
        for sec in doc.sections:
            sec.top_margin = Inches(0.8)
            sec.bottom_margin = Inches(0.8)
            sec.left_margin = Inches(0.8)
            sec.right_margin = Inches(0.8)

        p_title = doc.add_paragraph()
        p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r_title = p_title.add_run(f"BẢNG QUYẾT ĐỊNH NGHIỆP VỤ (DECISION TABLE)\n{self.table_id}: {self.title.upper()}")
        r_title.bold = True
        r_title.font.size = Pt(15)
        r_title.font.color.rgb = RGBColor(15, 23, 42)

        p_sub = doc.add_paragraph()
        p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r_sub = p_sub.add_run(f"Phân hệ: {self.domain} | Chuẩn BABOK v3 | Ngày xuất: {datetime.date.today().strftime('%d/%m/%Y')}")
        r_sub.italic = True
        r_sub.font.size = Pt(9.5)
        r_sub.font.color.rgb = RGBColor(100, 116, 139)

        if self.description:
            p_desc = doc.add_paragraph()
            p_desc.add_run(f"Mô tả nghiệp vụ: {self.description}").italic = True
            p_desc.paragraph_format.space_after = Pt(8)

        total_cols = 2 + len(self.rules)
        total_rows = 1 + len(self.conditions) + len(self.actions)

        table = doc.add_table(rows=total_rows, cols=total_cols)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        table.autofit = False

        col_w_type = Inches(1.1)
        col_w_name = Inches(2.6)
        col_w_rule = Inches(0.65)

        # Header Row
        hdr_row = table.rows[0]
        hdr_row._tr.get_or_add_trPr().append(parse_xml(f'<w:tblHeader {nsdecls("w")}/>'))
        hdr_row._tr.get_or_add_trPr().append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))

        hdr_cells = hdr_row.cells
        hdr_cells[0].width = col_w_type
        hdr_cells[1].width = col_w_name

        hdr_cells[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        r0 = hdr_cells[0].paragraphs[0].add_run("Phân Loại")
        r0.bold = True
        r0.font.color.rgb = RGBColor(255, 255, 255)
        r0.font.size = Pt(9)

        hdr_cells[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.LEFT
        r1 = hdr_cells[1].paragraphs[0].add_run("Điều Kiện / Hành Động")
        r1.bold = True
        r1.font.color.rgb = RGBColor(255, 255, 255)
        r1.font.size = Pt(9)

        for ci in range(2):
            tcPr = hdr_cells[ci]._tc.get_or_add_tcPr()
            tcPr.append(parse_xml(f'<w:shd {nsdecls("w")} w:fill="0F172A"/>'))
            tcPr.append(parse_xml(f'<w:vAlign {nsdecls("w")} w:val="center"/>'))

        for r_idx, rule in enumerate(self.rules, start=1):
            cell = hdr_cells[1 + r_idx]
            cell.width = col_w_rule
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            rr = p.add_run(rule.rule_id)
            rr.bold = True
            rr.font.color.rgb = RGBColor(255, 255, 255)
            rr.font.size = Pt(9)
            tcPr = cell._tc.get_or_add_tcPr()
            tcPr.append(parse_xml(f'<w:shd {nsdecls("w")} w:fill="1E293B"/>'))
            tcPr.append(parse_xml(f'<w:vAlign {nsdecls("w")} w:val="center"/>'))

        # Data Rows - Conditions
        row_idx = 1
        for c_idx, cond in enumerate(self.conditions, start=1):
            row = table.rows[row_idx]
            row._tr.get_or_add_trPr().append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))

            cells = row.cells
            cells[0].width = col_w_type
            cells[1].width = col_w_name

            p0 = cells[0].paragraphs[0]
            p0.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r_c = p0.add_run("ĐIỀU KIỆN" if c_idx == 1 else "")
            r_c.bold = True
            r_c.font.size = Pt(8.5)
            r_c.font.color.rgb = RGBColor(2, 132, 199)

            p1 = cells[1].paragraphs[0]
            r_cn = p1.add_run(cond.name)
            r_cn.font.size = Pt(8.5)
            r_cn.bold = True

            for ci in range(2):
                tcPr = cells[ci]._tc.get_or_add_tcPr()
                tcPr.append(parse_xml(f'<w:vAlign {nsdecls("w")} w:val="center"/>'))
                tcPr.append(parse_xml(f'<w:shd {nsdecls("w")} w:fill="F0F9FF"/>'))

            for r_idx, rule in enumerate(self.rules, start=1):
                cell = cells[1 + r_idx]
                cell.width = col_w_rule
                p = cell.paragraphs[0]
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                val = str(rule.conditions.get(cond.name, "-"))
                rv = p.add_run(val)
                rv.font.name = "Consolas"
                rv.font.size = Pt(9)
                rv.bold = True
                if val in ("Y", "YES", "TRUE", "X"):
                    rv.font.color.rgb = RGBColor(2, 132, 199)

                tcPr = cell._tc.get_or_add_tcPr()
                tcPr.append(parse_xml(f'<w:vAlign {nsdecls("w")} w:val="center"/>'))
                tcPr.append(parse_xml(f'<w:shd {nsdecls("w")} w:fill="F8FAFC"/>'))

            row_idx += 1

        # Data Rows - Actions
        for a_idx, act in enumerate(self.actions, start=1):
            row = table.rows[row_idx]
            row._tr.get_or_add_trPr().append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))

            cells = row.cells
            cells[0].width = col_w_type
            cells[1].width = col_w_name

            p0 = cells[0].paragraphs[0]
            p0.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r_a = p0.add_run("HÀNH ĐỘNG" if a_idx == 1 else "")
            r_a.bold = True
            r_a.font.size = Pt(8.5)
            r_a.font.color.rgb = RGBColor(22, 163, 74)

            p1 = cells[1].paragraphs[0]
            r_an = p1.add_run(act.name)
            r_an.font.size = Pt(8.5)
            r_an.bold = True

            for ci in range(2):
                tcPr = cells[ci]._tc.get_or_add_tcPr()
                tcPr.append(parse_xml(f'<w:vAlign {nsdecls("w")} w:val="center"/>'))
                tcPr.append(parse_xml(f'<w:shd {nsdecls("w")} w:fill="F0FDF4"/>'))

            for r_idx, rule in enumerate(self.rules, start=1):
                cell = cells[1 + r_idx]
                cell.width = col_w_rule
                p = cell.paragraphs[0]
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                val = str(rule.actions.get(act.name, "-"))
                rv = p.add_run(val)
                rv.font.name = "Consolas"
                rv.font.size = Pt(9)
                rv.bold = True
                if val in ("X", "YES", "EXECUTE", "DUYET"):
                    rv.font.color.rgb = RGBColor(22, 163, 74)

                tcPr = cell._tc.get_or_add_tcPr()
                tcPr.append(parse_xml(f'<w:vAlign {nsdecls("w")} w:val="center"/>'))
                tcPr.append(parse_xml(f'<w:shd {nsdecls("w")} w:fill="FAFAFA"/>'))

            row_idx += 1

        os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
        doc.save(output_path)
        return output_path

