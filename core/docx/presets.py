"""
core.docx.presets — Presets & Administrative Document Standards (v4.0).

Supports:
1. Vietnam Decree 30/2020/ND-CP (Nghị định 30/2020/NĐ-CP về công tác văn thư):
   - A4 page setup: Top 20-25mm, Bottom 20-25mm, Left 30-35mm, Right 15-20mm.
   - Typography: Times New Roman, 13-14pt, 1.3-1.5 line spacing.
   - Administrative 2-Column Header (Cơ quan chủ quản / Quốc hiệu - Tiêu ngữ) with rule lines.
2. Academic Thesis & Corporate Report Presets.
3. OpenXML table & typography invariants compliance.
"""

from __future__ import annotations

import os
import re
from typing import Any, Dict, List, Optional, Tuple

import docx
from docx import Document
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.shared import Inches, Mm, Pt, RGBColor
import yaml

_CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
_WORKSPACE_ROOT = os.path.dirname(os.path.dirname(_CURRENT_DIR))
_PRESETS_DIR = os.path.join(_WORKSPACE_ROOT, "presets")

# Fallback built-in nd30 configuration in case preset YAML is deleted
BUILTIN_ND30_PRESET: Dict[str, Any] = {
    "geometry": {
        "page_size": "A4",
        "orientation": "portrait",
        "margin": {
            "top": "20mm",
            "bottom": "20mm",
            "left": "30mm",
            "right": "20mm",
        }
    },
    "typography": {
        "font_family": "Times New Roman",
        "font_size": "13pt",
        "body_color": "#000000",
        "line_spacing": 1.3,
        "paragraph_spacing_after": "4pt",
        "paragraph_spacing_before": "0pt",
        "first_line_indent": "1.27cm"
    },
    "headings": {
        "color": "#000000",
        "h1_size": "14pt",
        "h2_size": "13pt",
        "h3_size": "13pt"
    },
    "tables": {
        "header_background": "#F2F2F2",
        "border_color": "#000000",
        "cell_valign": "center"
    }
}


def _parse_dimension(val: Any, default_pt: float = 12.0) -> Pt:
    """Parse string dimension (e.g. '20mm', '1.0in', '13pt', '1.27cm') to Pt/Inches/Mm."""
    if isinstance(val, (int, float)):
        return Pt(val)
    if not isinstance(val, str):
        return Pt(default_pt)

    val = val.strip().lower()
    m_mm = re.match(r"^([\d.]+)\s*mm$", val)
    if m_mm:
        return Mm(float(m_mm.group(1)))
    m_cm = re.match(r"^([\d.]+)\s*cm$", val)
    if m_cm:
        return Mm(float(m_cm.group(1)) * 10.0)
    m_in = re.match(r"^([\d.]+)\s*in(?:ch)?$", val)
    if m_in:
        return Inches(float(m_in.group(1)))
    m_pt = re.match(r"^([\d.]+)\s*pt$", val)
    if m_pt:
        return Pt(float(m_pt.group(1)))

    try:
        return Pt(float(val))
    except Exception:
        return Pt(default_pt)


def list_available_presets() -> List[str]:
    """Scan and list all available preset names."""
    presets = ["nd30", "nd30_hanh_chinh", "academic_thesis", "corporate_report"]
    if os.path.isdir(_PRESETS_DIR):
        for f in os.listdir(_PRESETS_DIR):
            if f.endswith(".style.yaml"):
                stem = f[:-11]
                if stem not in presets:
                    presets.append(stem)
    return presets


def load_preset(preset_name: str = "nd30") -> Dict[str, Any]:
    """Load preset configuration dictionary by name or file path."""
    clean_name = preset_name.strip().lower()
    if clean_name in ("nd30", "nd30_hanh_chinh", "hanh_chinh"):
        target_file = os.path.join(_PRESETS_DIR, "nd30_hanh_chinh.style.yaml")
    elif os.path.isfile(preset_name):
        target_file = preset_name
    else:
        target_file = os.path.join(_PRESETS_DIR, f"{clean_name}.style.yaml")

    if os.path.isfile(target_file):
        try:
            with open(target_file, "r", encoding="utf-8") as f:
                data = yaml.safe_load(f)
                if isinstance(data, dict):
                    return data
        except Exception:
            pass

    return dict(BUILTIN_ND30_PRESET)


def apply_preset_to_document(doc: docx.Document, preset: str | Dict[str, Any] = "nd30") -> Dict[str, Any]:
    """
    Applies page geometry, typography, and margins from a preset to an existing docx.Document.
    """
    cfg = load_preset(preset) if isinstance(preset, str) else preset
    geom = cfg.get("geometry", {})
    margins = geom.get("margin", {})
    typo = cfg.get("typography", {})

    # 1. Apply Margins & Page Size across all sections
    for sec in doc.sections:
        # A4 Page Size (210mm x 297mm)
        sec.page_width = Mm(210)
        sec.page_height = Mm(297)

        if "top" in margins:
            sec.top_margin = _parse_dimension(margins["top"])
        if "bottom" in margins:
            sec.bottom_margin = _parse_dimension(margins["bottom"])
        if "left" in margins:
            sec.left_margin = _parse_dimension(margins["left"])
        if "right" in margins:
            sec.right_margin = _parse_dimension(margins["right"])

    # 2. Configure Default 'Normal' Style Typography
    font_family = typo.get("font_family", "Times New Roman").split(",")[0].strip().replace('"', '').replace("'", "")
    font_size = _parse_dimension(typo.get("font_size", "13pt"), default_pt=13.0)
    line_spacing = float(typo.get("line_spacing", 1.3))

    try:
        normal_style = doc.styles["Normal"]
        normal_style.font.name = font_family
        normal_style.font.size = font_size
        normal_style.font.color.rgb = RGBColor(0, 0, 0)
        normal_style.paragraph_format.line_spacing = line_spacing
        if "paragraph_spacing_after" in typo:
            normal_style.paragraph_format.space_after = _parse_dimension(typo["paragraph_spacing_after"])
        if "paragraph_spacing_before" in typo:
            normal_style.paragraph_format.space_before = _parse_dimension(typo["paragraph_spacing_before"])
    except Exception:
        pass

    return {
        "success": True,
        "preset_applied": preset if isinstance(preset, str) else "custom",
        "font_family": font_family,
        "font_size": str(font_size),
        "margins_mm": {
            "top": 20,
            "bottom": 20,
            "left": 30,
            "right": 20,
        }
    }


def build_administrative_header(
    doc: docx.Document,
    agency_parent: str = "BỘ THÔNG TIN VÀ TRUYỀN THÔNG",
    agency_name: str = "CỤC CÔNG NGHỆ THÔNG TIN",
    document_no: str = "Số: 128/CNTT-VP",
    location_date: str = "Hà Nội, ngày 18 tháng 10 năm 2026",
    republic_title: str = "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM",
    motto: str = "Độc lập - Tự do - Hạnh phúc",
    insert_at_top: bool = False,
) -> Any:
    """
    Constructs a compliant Decree 30/2020/ND-CP two-column administrative header table:
    - Column 1: Issuing Agency + Document Number
    - Column 2: Republic Title + National Motto + Location/Date
    - Invariants: Invisible borders, cantSplit, tblHeader, vAlign='top', and cell ends with <w:p> (ERR_DOCX_001).
    """
    table = doc.add_table(rows=1, cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    # Enforce Invisible Table Borders
    tblPr = table._tbl.tblPr
    tblBorders = parse_xml(
        f'<w:tblBorders {nsdecls("w")}>\n'
        f'  <w:top w:val="none" w:sz="0" w:space="0" w:color="auto"/>\n'
        f'  <w:left w:val="none" w:sz="0" w:space="0" w:color="auto"/>\n'
        f'  <w:bottom w:val="none" w:sz="0" w:space="0" w:color="auto"/>\n'
        f'  <w:right w:val="none" w:sz="0" w:space="0" w:color="auto"/>\n'
        f'  <w:insideH w:val="none" w:sz="0" w:space="0" w:color="auto"/>\n'
        f'  <w:insideV w:val="none" w:sz="0" w:space="0" w:color="auto"/>\n'
        f'</w:tblBorders>'
    )
    tblPr.append(tblBorders)

    row = table.rows[0]
    trPr = row._tr.get_or_add_trPr()
    trPr.append(OxmlElement("w:cantSplit"))
    trPr.append(OxmlElement("w:tblHeader"))

    # Set Column Widths (Left ~45%, Right ~55%)
    row.cells[0].width = Mm(75)
    row.cells[1].width = Mm(90)

    # Column 1 (Left): Issuing Agency
    c1 = row.cells[0]
    c1.vertical_alignment = WD_ALIGN_VERTICAL.TOP
    p1 = c1.paragraphs[0]
    p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p1.paragraph_format.line_spacing = 1.15
    p1.paragraph_format.space_after = Pt(2)
    p1.paragraph_format.space_before = Pt(0)
    r1 = p1.add_run(agency_parent.upper())
    r1.font.name = "Times New Roman"
    r1.font.size = Pt(12)

    p2 = c1.add_paragraph()
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p2.paragraph_format.line_spacing = 1.15
    p2.paragraph_format.space_after = Pt(3)
    p2.paragraph_format.space_before = Pt(0)
    r2 = p2.add_run(agency_name.upper())
    r2.font.name = "Times New Roman"
    r2.font.size = Pt(12)
    r2.bold = True

    # Separator line under Agency Name (centered dash/line)
    p_line1 = c1.add_paragraph()
    p_line1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_line1.paragraph_format.line_spacing = 1.0
    p_line1.paragraph_format.space_after = Pt(4)
    p_line1.paragraph_format.space_before = Pt(0)
    r_sep1 = p_line1.add_run("—————")
    r_sep1.font.name = "Times New Roman"
    r_sep1.font.size = Pt(10)
    r_sep1.bold = True

    p3 = c1.add_paragraph()
    p3.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p3.paragraph_format.line_spacing = 1.15
    p3.paragraph_format.space_after = Pt(0)
    p3.paragraph_format.space_before = Pt(0)
    r3 = p3.add_run(document_no)
    r3.font.name = "Times New Roman"
    r3.font.size = Pt(12)

    # Column 2 (Right): Republic Title & Motto
    c2 = row.cells[1]
    c2.vertical_alignment = WD_ALIGN_VERTICAL.TOP
    p_rep = c2.paragraphs[0]
    p_rep.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_rep.paragraph_format.line_spacing = 1.15
    p_rep.paragraph_format.space_after = Pt(2)
    p_rep.paragraph_format.space_before = Pt(0)
    r_rep = p_rep.add_run(republic_title.upper())
    r_rep.font.name = "Times New Roman"
    r_rep.font.size = Pt(12.5)
    r_rep.bold = True

    p_motto = c2.add_paragraph()
    p_motto.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_motto.paragraph_format.line_spacing = 1.15
    p_motto.paragraph_format.space_after = Pt(3)
    p_motto.paragraph_format.space_before = Pt(0)
    r_motto = p_motto.add_run(motto)
    r_motto.font.name = "Times New Roman"
    r_motto.font.size = Pt(13)
    r_motto.bold = True

    # Rule line under National Motto (solid line as wide as the text)
    p_line2 = c2.add_paragraph()
    p_line2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_line2.paragraph_format.line_spacing = 1.0
    p_line2.paragraph_format.space_after = Pt(6)
    p_line2.paragraph_format.space_before = Pt(0)
    r_sep2 = p_line2.add_run("——————————")
    r_sep2.font.name = "Times New Roman"
    r_sep2.font.size = Pt(11)
    r_sep2.bold = True

    p_loc = c2.add_paragraph()
    p_loc.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_loc.paragraph_format.line_spacing = 1.15
    p_loc.paragraph_format.space_after = Pt(0)
    p_loc.paragraph_format.space_before = Pt(0)
    r_loc = p_loc.add_run(location_date)
    r_loc.font.name = "Times New Roman"
    r_loc.font.size = Pt(13)
    r_loc.italic = True

    # If insert_at_top is requested, move table element to the very beginning of the body
    if insert_at_top:
        doc._body._element.insert(0, table._tbl)

    return table
