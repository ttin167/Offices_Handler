"""
core.docx.design_system — Centralized Enterprise Document Design System & Premium Styler.

Implements:
1. Multi-Theme Registry:
   - 'modern_tech' (Default): Clean Slate/Indigo theme (Inter / Segoe UI / Aptos, #0F172A, #2563EB, #E2E8F0).
   - 'executive_classic': Formal Corporate Report (Times New Roman, #0F294A, #881337).
   - 'nghidinh30': Vietnam Administrative Standard (Decree 30/2020/ND-CP).
2. Executive Cover Page Builder (Modern Bauhaus accent bar, metadata grid card, clean typography).
3. Native Dynamic Table of Contents (Word TOC field with F9 update capability).
4. Dual Header/Footer Injector (Right-aligned title header + 'Trang X / Y' page numbering footer).
5. High-Impact Callout Cards (NOTE, TIP, WARNING, IMPORTANT) with colored accent borders and soft background.
6. Fenced Code Block Cards with language badges and Consolas typography.
7. Modern Clean Studio Table Engine (6pt/8pt cell padding, ultra-thin borders, intelligent data auto-alignment).
8. Template Token Formatters (Currency, Dates, Number formatting).
"""

from __future__ import annotations

import datetime
import html
import re
from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional, Tuple, Union

import docx
from docx import Document
from docx.document import Document as DocumentType
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.shared import Cm, Inches, Mm, Pt, RGBColor

NS_W = nsdecls("w")


# ---------------------------------------------------------------------------
# 1. Theme Data Structure & Registry
# ---------------------------------------------------------------------------

@dataclass
class DocxTheme:
    name: str
    font_primary: str = "Segoe UI"
    font_heading: str = "Segoe UI"
    font_code: str = "Consolas"
    
    # Color Palette (Hex without #)
    color_primary: str = "0F172A"       # Deep Slate / Dark Title
    color_accent: str = "2563EB"        # Royal Blue / Interactive
    color_body: str = "334155"          # Charcoal Text
    color_muted: str = "64748B"         # Muted Gray
    
    # Table Styling
    table_header_fill: str = "0F172A"   # Dark Navy / Slate Header
    table_header_text: str = "FFFFFF"   # Pure White Text
    table_zebra_fill: str = "F8FAFC"    # Soft Light Row
    table_border_color: str = "E2E8F0"  # Ultra-thin Clean Border
    table_pad_top_pt: float = 6.0
    table_pad_bottom_pt: float = 6.0
    table_pad_left_pt: float = 8.0
    table_pad_right_pt: float = 8.0

    # Margins (Centimeters)
    margin_top_cm: float = 2.2
    margin_bottom_cm: float = 2.2
    margin_left_cm: float = 2.5
    margin_right_cm: float = 2.0

    # Line & Paragraph Spacing
    line_spacing: float = 1.18
    space_after_pt: float = 4.0

    def hex_to_rgb(self, hex_val: str) -> RGBColor:
        clean = hex_val.lstrip("#")
        if len(clean) == 6:
            return RGBColor(int(clean[0:2], 16), int(clean[2:4], 16), int(clean[4:6], 16))
        return RGBColor(15, 23, 42)


THEME_MODERN_TECH = DocxTheme(
    name="modern_tech",
    font_primary="Segoe UI",
    font_heading="Segoe UI",
    font_code="Consolas",
    color_primary="0F172A",
    color_accent="2563EB",
    color_body="1E293B",
    color_muted="64748B",
    table_header_fill="0F172A",
    table_header_text="FFFFFF",
    table_zebra_fill="F8FAFC",
    table_border_color="E2E8F0",
    table_pad_top_pt=6.0,
    table_pad_bottom_pt=6.0,
    table_pad_left_pt=8.0,
    table_pad_right_pt=8.0,
    margin_top_cm=2.2,
    margin_bottom_cm=2.2,
    margin_left_cm=2.5,
    margin_right_cm=2.0,
    line_spacing=1.18,
    space_after_pt=4.0,
)

THEME_EXECUTIVE_CLASSIC = DocxTheme(
    name="executive_classic",
    font_primary="Times New Roman",
    font_heading="Times New Roman",
    font_code="Courier New",
    color_primary="0F294A",
    color_accent="881337",
    color_body="000000",
    color_muted="475569",
    table_header_fill="0F294A",
    table_header_text="FFFFFF",
    table_zebra_fill="F1F5F9",
    table_border_color="CBD5E1",
    table_pad_top_pt=5.0,
    table_pad_bottom_pt=5.0,
    table_pad_left_pt=7.0,
    table_pad_right_pt=7.0,
    margin_top_cm=2.5,
    margin_bottom_cm=2.5,
    margin_left_cm=3.0,
    margin_right_cm=2.0,
    line_spacing=1.25,
    space_after_pt=5.0,
)

THEME_NGHIDINH30 = DocxTheme(
    name="nghidinh30",
    font_primary="Times New Roman",
    font_heading="Times New Roman",
    font_code="Courier New",
    color_primary="000000",
    color_accent="000000",
    color_body="000000",
    color_muted="333333",
    table_header_fill="F2F2F2",
    table_header_text="000000",
    table_zebra_fill="FFFFFF",
    table_border_color="000000",
    table_pad_top_pt=4.0,
    table_pad_bottom_pt=4.0,
    table_pad_left_pt=6.0,
    table_pad_right_pt=6.0,
    margin_top_cm=2.0,
    margin_bottom_cm=2.0,
    margin_left_cm=3.0,
    margin_right_cm=1.8,
    line_spacing=1.35,
    space_after_pt=6.0,
)

THEME_REGISTRY: Dict[str, DocxTheme] = {
    "modern_tech": THEME_MODERN_TECH,
    "modern": THEME_MODERN_TECH,
    "default": THEME_MODERN_TECH,
    "executive_classic": THEME_EXECUTIVE_CLASSIC,
    "executive": THEME_EXECUTIVE_CLASSIC,
    "classic": THEME_EXECUTIVE_CLASSIC,
    "nghidinh30": THEME_NGHIDINH30,
    "nd30": THEME_NGHIDINH30,
}


def get_theme(name: Optional[str] = None) -> DocxTheme:
    """Retrieve theme by key name, falling back to modern_tech."""
    if not name:
        return THEME_MODERN_TECH
    key = name.strip().lower().replace("-", "_")
    return THEME_REGISTRY.get(key, THEME_MODERN_TECH)


# ---------------------------------------------------------------------------
# 2. Executive Cover Page Builder
# ---------------------------------------------------------------------------

def build_cover_page(
    doc: DocumentType,
    title: str,
    subtitle: str = "",
    author: str = "Lead Solution Architect",
    org: str = "KHỐI CÔNG NGHỆ THÔNG TIN & CHUYỂN ĐỔI SỐ",
    date_str: Optional[str] = None,
    version_str: str = "v1.0 (Enterprise Specification)",
    theme: Optional[DocxTheme] = None,
) -> None:
    """
    Renders a publication-grade Executive Cover Page with geometric accent band,
    spacious title typography, and a structured metadata card.
    """
    th = theme or THEME_MODERN_TECH
    today = date_str or datetime.date.today().strftime("%d/%m/%Y")

    # Top accent band: Decorative colored divider bar
    p_band = doc.add_paragraph()
    p_band.paragraph_format.space_before = Pt(12)
    p_band.paragraph_format.space_after = Pt(28)
    pPr = p_band._p.get_or_add_pPr()
    pBdr = parse_xml(
        f'<w:pBdr {NS_W}>'
        f'  <w:bottom w:val="single" w:sz="36" w:space="1" w:color="{th.color_accent}"/>'
        f'</w:pBdr>'
    )
    pPr.append(pBdr)

    # Organization / Sub-header
    if org:
        p_org = doc.add_paragraph()
        p_org.paragraph_format.space_after = Pt(36)
        r_org = p_org.add_run(org.upper())
        r_org.font.name = th.font_heading
        r_org.font.size = Pt(10.5)
        r_org.bold = True
        r_org.font.color.rgb = th.hex_to_rgb(th.color_muted)

    # Title
    p_title = doc.add_paragraph()
    p_title.paragraph_format.space_after = Pt(12)
    p_title.paragraph_format.line_spacing = 1.12
    r_title = p_title.add_run(title.upper())
    r_title.font.name = th.font_heading
    r_title.font.size = Pt(24)
    r_title.bold = True
    r_title.font.color.rgb = th.hex_to_rgb(th.color_primary)

    # Subtitle
    if subtitle:
        p_sub = doc.add_paragraph()
        p_sub.paragraph_format.space_after = Pt(48)
        r_sub = p_sub.add_run(subtitle)
        r_sub.font.name = th.font_primary
        r_sub.font.size = Pt(13)
        r_sub.italic = True
        r_sub.font.color.rgb = th.hex_to_rgb(th.color_accent)
    else:
        doc.add_paragraph().paragraph_format.space_after = Pt(36)

    # Spacer
    for _ in range(3):
        doc.add_paragraph()

    # Metadata Grid Box (Clean single-cell card with left accent bar)
    t_card = doc.add_table(rows=1, cols=1)
    t_card.alignment = WD_TABLE_ALIGNMENT.CENTER
    t_card.autofit = False

    trPr_card = t_card.rows[0]._tr.get_or_add_trPr()
    trPr_card.append(parse_xml(f'<w:tblHeader {NS_W}/>'))
    trPr_card.append(parse_xml(f'<w:cantSplit {NS_W}/>'))

    cell_card = t_card.cell(0, 0)
    cell_card.width = Inches(6.5)

    tcPr = cell_card._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {NS_W} w:fill="{th.table_zebra_fill}"/>')
    tcBorders = parse_xml(
        f'<w:tcBorders {NS_W}>'
        f'  <w:left w:val="single" w:sz="24" w:space="0" w:color="{th.color_accent}"/>'
        f'  <w:top w:val="single" w:sz="4" w:space="0" w:color="{th.table_border_color}"/>'
        f'  <w:bottom w:val="single" w:sz="4" w:space="0" w:color="{th.table_border_color}"/>'
        f'  <w:right w:val="single" w:sz="4" w:space="0" w:color="{th.table_border_color}"/>'
        f'</w:tcBorders>'
    )
    tcMar = parse_xml(
        f'<w:tcMar {NS_W}>'
        f'  <w:top w:w="240" w:type="dxa"/>'
        f'  <w:bottom w:w="240" w:type="dxa"/>'
        f'  <w:left w:w="300" w:type="dxa"/>'
        f'  <w:right w:w="300" w:type="dxa"/>'
        f'</w:tcMar>'
    )
    tcPr.append(shd)
    tcPr.append(tcBorders)
    tcPr.append(tcMar)

    p_c0 = cell_card.paragraphs[0]
    p_c0.paragraph_format.space_before = Pt(4)
    p_c0.paragraph_format.space_after = Pt(2)
    r_hdr = p_c0.add_run("THÔNG TIN HỒ SƠ TÀI LIỆU")
    r_hdr.bold = True
    r_hdr.font.size = Pt(10)
    r_hdr.font.color.rgb = th.hex_to_rgb(th.color_primary)

    meta_items = [
        ("Tác giả / Chủ biên:", author),
        ("Phiên bản đặc tả:", version_str),
        ("Ngày phát hành:", today),
        ("Môi trường biên dịch:", "Antigravity High-Fidelity Doc Engine"),
    ]
    for label, val in meta_items:
        p_row = cell_card.add_paragraph()
        p_row.paragraph_format.space_before = Pt(1)
        p_row.paragraph_format.space_after = Pt(2)
        r_l = p_row.add_run(f"• {label} ")
        r_l.bold = True
        r_l.font.size = Pt(9.5)
        r_l.font.color.rgb = th.hex_to_rgb(th.color_muted)
        r_v = p_row.add_run(val)
        r_v.font.size = Pt(9.5)
        r_v.font.color.rgb = th.hex_to_rgb(th.color_body)

    # End with clean page break
    doc.add_page_break()


# ---------------------------------------------------------------------------
# 3. Native Dynamic Table of Contents (Word F9 TOC Field)
# ---------------------------------------------------------------------------

def inject_native_toc(
    doc: DocumentType,
    title: str = "MỤC LỤC TÀI LIỆU",
    max_level: int = 3,
    theme: Optional[DocxTheme] = None,
) -> None:
    """
    Injects a native Word Table of Contents field code:
    <w:fldSimple w:instr="TOC \\o &quot;1-3&quot; \\h \\z \\u"/>
    Word recalculates exact page numbers upon opening or when pressing F9.
    """
    th = theme or THEME_MODERN_TECH

    p_toc_title = doc.add_paragraph()
    p_toc_title.paragraph_format.space_before = Pt(14)
    p_toc_title.paragraph_format.space_after = Pt(10)
    r_t = p_toc_title.add_run(title.upper())
    r_t.font.name = th.font_heading
    r_t.font.size = Pt(16)
    r_t.bold = True
    r_t.font.color.rgb = th.hex_to_rgb(th.color_primary)

    p_field = doc.add_paragraph()
    p_field.paragraph_format.space_after = Pt(16)
    fldSimple = OxmlElement("w:fldSimple")
    fldSimple.set(qn("w:instr"), f'TOC \\o "1-{max_level}" \\h \\z \\u')

    # Fallback initial prompt text inside field
    r_fld = OxmlElement("w:r")
    rPr = OxmlElement("w:rPr")
    i_el = OxmlElement("w:i")
    rPr.append(i_el)
    color_el = OxmlElement("w:color")
    color_el.set(qn("w:val"), th.color_muted)
    rPr.append(color_el)
    t_el = OxmlElement("w:t")
    t_el.text = "Nhấp chuột phải hoặc bấm F9 để cập nhật toàn bộ mục lục..."
    r_fld.append(rPr)
    r_fld.append(t_el)
    fldSimple.append(r_fld)

    p_field._p.append(fldSimple)
    doc.add_page_break()


# ---------------------------------------------------------------------------
# 4. Dual Header & Footer with 'Trang X / Y' Native Fields
# ---------------------------------------------------------------------------

def setup_header_footer(
    doc: DocumentType,
    header_title: str = "",
    classification: str = "TÀI LIỆU NỘI BỘ",
    theme: Optional[DocxTheme] = None,
    confidentiality: Optional[str] = None,
) -> None:
    """
    Configures sections with:
    - Right-aligned document title in header.
    - Footer with Left classification tag + Right 'Trang {PAGE} / {NUMPAGES}' field.
    """
    if confidentiality is not None:
        classification = confidentiality
    th = theme or THEME_MODERN_TECH

    for sec in doc.sections:
        # 1. Header
        if header_title:
            hp = sec.header.paragraphs[0] if sec.header.paragraphs else sec.header.add_paragraph()
            hp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            hp.text = ""
            r_hdr = hp.add_run(header_title)
            r_hdr.font.name = th.font_primary
            r_hdr.font.size = Pt(8.5)
            r_hdr.font.color.rgb = th.hex_to_rgb(th.color_muted)

        # 2. Footer
        fp = sec.footer.paragraphs[0] if sec.footer.paragraphs else sec.footer.add_paragraph()
        fp.text = ""
        fp.paragraph_format.space_before = Pt(4)
        fp.paragraph_format.space_after = Pt(0)

        # Tab stop at right margin for 2-column footer (Left: Class, Right: Page X/Y)
        try:
            pw = sec.page_width or Inches(8.5)
            lm = sec.left_margin or Inches(1.0)
            rm = sec.right_margin or Inches(1.0)
            avail_w_twips = int((int(pw) - int(lm) - int(rm)) / 635)
        except Exception:
            avail_w_twips = 9360

        pPr = fp._p.get_or_add_pPr()
        tabs_el = parse_xml(
            f'<w:tabs {NS_W}>'
            f'  <w:tab w:val="right" w:pos="{avail_w_twips}"/>'
            f'</w:tabs>'
        )
        pPr.append(tabs_el)

        # Left label: Classification
        r_left = fp.add_run(classification.upper())
        r_left.font.name = th.font_primary
        r_left.font.size = Pt(8.5)
        r_left.font.bold = True
        r_left.font.color.rgb = th.hex_to_rgb(th.color_muted)

        # Tab separator
        r_tab = fp.add_run("\t")

        # Right label: Trang {PAGE} / {NUMPAGES}
        r_pg_label = fp.add_run("Trang ")
        r_pg_label.font.name = th.font_primary
        r_pg_label.font.size = Pt(8.5)
        r_pg_label.font.color.rgb = th.hex_to_rgb(th.color_muted)

        # Dynamic PAGE field
        fld_page = OxmlElement("w:fldSimple")
        fld_page.set(qn("w:instr"), "PAGE")
        r_p = OxmlElement("w:r")
        t_p = OxmlElement("w:t")
        t_p.text = "1"
        r_p.append(t_p)
        fld_page.append(r_p)
        fp._p.append(fld_page)

        r_sep = fp.add_run(" / ")
        r_sep.font.name = th.font_primary
        r_sep.font.size = Pt(8.5)
        r_sep.font.color.rgb = th.hex_to_rgb(th.color_muted)

        # Dynamic NUMPAGES field
        fld_numpages = OxmlElement("w:fldSimple")
        fld_numpages.set(qn("w:instr"), "NUMPAGES")
        r_np = OxmlElement("w:r")
        t_np = OxmlElement("w:t")
        t_np.text = "1"
        r_np.append(t_np)
        fld_numpages.append(r_np)
        fp._p.append(fld_numpages)


# ---------------------------------------------------------------------------
# 5. High-Impact Callout Cards (NOTE, TIP, WARNING, IMPORTANT)
# ---------------------------------------------------------------------------

CALLOUT_SPECS: Dict[str, Dict[str, str]] = {
    "NOTE": {"border": "0284C7", "fill": "F0F9FF", "icon": "", "label": "LƯU Ý / GHI CHÚ"},
    "INFO": {"border": "0284C7", "fill": "F0F9FF", "icon": "", "label": "THÔNG TIN"},
    "TIP": {"border": "16A34A", "fill": "F0FDF4", "icon": "💡", "label": "MẸO HỮU ÍCH"},
    "WARNING": {"border": "D97706", "fill": "FFFBEB", "icon": "⚠", "label": "CẢNH BÁO"},
    "IMPORTANT": {"border": "DC2626", "fill": "FEF2F2", "icon": "🛑", "label": "QUAN TRỌNG"},
    "CAUTION": {"border": "DC2626", "fill": "FEF2F2", "icon": "🛑", "label": "CHÚ Ý NGUY HIỂM"},
}


def render_inline_markdown_runs(
    paragraph: Any,
    text: str,
    font_name: str = "Segoe UI",
    font_size_pt: float = 9.5,
    color_rgb: Optional[RGBColor] = None,
) -> None:
    """
    Renders inline markdown (bold **...**, italic *...*, code `...`) into docx runs.
    Eliminates raw markdown artifacts from callouts, table cells, and cards.
    """
    if not text:
        return

    # Split by inline code blocks `...` first
    code_parts = re.split(r"(`[^`]+`)", text)
    for c_part in code_parts:
        if not c_part:
            continue
        if c_part.startswith("`") and c_part.endswith("`") and len(c_part) >= 2:
            r = paragraph.add_run(c_part[1:-1])
            r.font.name = "Consolas"
            r.font.size = Pt(font_size_pt)
            if color_rgb:
                r.font.color.rgb = color_rgb
            continue

        # Split by bold **...** and italic *...*
        tokens = re.split(r"(\*\*.*?\*\*|\*[^*\n]+?\*)", c_part)
        for tok in tokens:
            if not tok:
                continue
            if tok.startswith("**") and tok.endswith("**") and len(tok) >= 4:
                r = paragraph.add_run(tok[2:-2])
                r.bold = True
                r.font.name = font_name
                r.font.size = Pt(font_size_pt)
                if color_rgb:
                    r.font.color.rgb = color_rgb
            elif tok.startswith("*") and tok.endswith("*") and len(tok) >= 2:
                r = paragraph.add_run(tok[1:-1])
                r.italic = True
                r.font.name = font_name
                r.font.size = Pt(font_size_pt)
                if color_rgb:
                    r.font.color.rgb = color_rgb
            else:
                r = paragraph.add_run(tok)
                r.font.name = font_name
                r.font.size = Pt(font_size_pt)
                if color_rgb:
                    r.font.color.rgb = color_rgb


def build_callout_box(
    doc: DocumentType,
    callout_type: str,
    content: Union[str, List[str]],
    title: Optional[str] = None,
    theme: Optional[DocxTheme] = None,
    render_fn: Optional[Any] = None,
) -> Any:
    """
    Renders an eye-catching callout card with a colored accent border and soft pastel fill.
    Supports inline Markdown formatting (bold, italic, code) in content lines.
    """
    th = theme or THEME_MODERN_TECH
    c_type = callout_type.strip().upper().replace("[!", "").replace("]", "").strip()
    spec = CALLOUT_SPECS.get(c_type, CALLOUT_SPECS["NOTE"])

    tbl = doc.add_table(rows=1, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl.autofit = False

    trPr_box = tbl.rows[0]._tr.get_or_add_trPr()
    trPr_box.append(parse_xml(f'<w:tblHeader {NS_W}/>'))
    trPr_box.append(parse_xml(f'<w:cantSplit {NS_W}/>'))

    cell = tbl.cell(0, 0)
    cell.width = Inches(6.5)

    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {NS_W} w:fill="{spec["fill"]}"/>')
    tcBorders = parse_xml(
        f'<w:tcBorders {NS_W}>'
        f'  <w:left w:val="single" w:sz="24" w:space="0" w:color="{spec["border"]}"/>'
        f'  <w:top w:val="none"/>'
        f'  <w:right w:val="none"/>'
        f'  <w:bottom w:val="none"/>'
        f'</w:tcBorders>'
    )
    tcMar = parse_xml(
        f'<w:tcMar {NS_W}>'
        f'  <w:top w:w="160" w:type="dxa"/>'
        f'  <w:bottom w:w="160" w:type="dxa"/>'
        f'  <w:left w:w="240" w:type="dxa"/>'
        f'  <w:right w:w="240" w:type="dxa"/>'
        f'</w:tcMar>'
    )
    tcPr.append(shd)
    tcPr.append(tcBorders)
    tcPr.append(tcMar)

    p0 = cell.paragraphs[0]
    p0.paragraph_format.space_before = Pt(3)
    p0.paragraph_format.space_after = Pt(2)
    header_label = title or spec["label"]
    icon_val = spec.get("icon", "").strip()
    if icon_val:
        hdr_text = f"{icon_val}  {header_label}"
    elif header_label.startswith("[") and header_label.endswith("]"):
        hdr_text = header_label
    else:
        hdr_text = f"[{header_label}]"
    r_hdr = p0.add_run(hdr_text)
    r_hdr.bold = True
    r_hdr.font.name = th.font_heading
    r_hdr.font.size = Pt(9.5)
    r_hdr.font.color.rgb = th.hex_to_rgb(spec["border"])

    lines = [content] if isinstance(content, str) else content
    for line in lines:
        stripped = line.strip()
        if not stripped:
            continue
        p = cell.add_paragraph()
        p.paragraph_format.space_before = Pt(1)
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.line_spacing = 1.15

        if render_fn is not None:
            render_fn(p, stripped)
        else:
            render_inline_markdown_runs(
                p,
                stripped,
                font_name=th.font_primary,
                font_size_pt=9.5,
                color_rgb=th.hex_to_rgb(th.color_body),
            )

    # Empty paragraph spacer after table
    p_spacer = doc.add_paragraph()
    p_spacer.paragraph_format.space_after = Pt(6)
    return tbl


# ---------------------------------------------------------------------------
# 6. Fenced Code Block Card Generator
# ---------------------------------------------------------------------------

def build_code_card(
    doc: DocumentType,
    code_text: str = "",
    language: str = "",
    theme: Optional[DocxTheme] = None,
    code: Optional[str] = None,
    lang: Optional[str] = None,
) -> Any:
    """
    Renders code in a stylized single-cell card with subtle background,
    language badge pill, and Consolas typography.
    """
    th = theme or THEME_MODERN_TECH
    actual_code = code if code is not None else code_text
    actual_lang = lang if lang is not None else language

    tbl = doc.add_table(rows=1, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl.autofit = False

    trPr_box = tbl.rows[0]._tr.get_or_add_trPr()
    trPr_box.append(parse_xml(f'<w:tblHeader {NS_W}/>'))
    trPr_box.append(parse_xml(f'<w:cantSplit {NS_W}/>'))

    cell = tbl.cell(0, 0)
    cell.width = Inches(6.5)

    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {NS_W} w:fill="F8FAFC"/>')
    tcBorders = parse_xml(
        f'<w:tcBorders {NS_W}>'
        f'  <w:left w:val="single" w:sz="6" w:space="0" w:color="{th.table_border_color}"/>'
        f'  <w:top w:val="single" w:sz="6" w:space="0" w:color="{th.table_border_color}"/>'
        f'  <w:right w:val="single" w:sz="6" w:space="0" w:color="{th.table_border_color}"/>'
        f'  <w:bottom w:val="single" w:sz="6" w:space="0" w:color="{th.table_border_color}"/>'
        f'</w:tcBorders>'
    )
    tcMar = parse_xml(
        f'<w:tcMar {NS_W}>'
        f'  <w:top w:w="160" w:type="dxa"/>'
        f'  <w:bottom w:w="160" w:type="dxa"/>'
        f'  <w:left w:w="220" w:type="dxa"/>'
        f'  <w:right w:w="220" w:type="dxa"/>'
        f'</w:tcMar>'
    )
    tcPr.append(shd)
    tcPr.append(tcBorders)
    tcPr.append(tcMar)

    p0 = cell.paragraphs[0]
    p0.paragraph_format.space_before = Pt(2)
    p0.paragraph_format.space_after = Pt(4)

    # Language badge pill
    lang_tag = actual_lang.strip().upper() if actual_lang else "CODE"
    r_lang = p0.add_run(f"[{lang_tag}]")
    r_lang.bold = True
    r_lang.font.name = th.font_code
    r_lang.font.size = Pt(8.0)
    r_lang.font.color.rgb = th.hex_to_rgb(th.color_accent)

    # Code lines
    for line in actual_code.splitlines():
        p_code = cell.add_paragraph()
        p_code.paragraph_format.space_before = Pt(0)
        p_code.paragraph_format.space_after = Pt(1)
        p_code.paragraph_format.line_spacing = 1.05
        r_c = p_code.add_run(line if line else " ")
        r_c.font.name = th.font_code
        r_c.font.size = Pt(9.0)
        r_c.font.color.rgb = th.hex_to_rgb(th.color_primary)

    p_spacer = doc.add_paragraph()
    p_spacer.paragraph_format.space_after = Pt(6)
    return tbl


# ---------------------------------------------------------------------------
# 7. Modern Clean Studio Table Engine
# ---------------------------------------------------------------------------

_NUMBER_REGEX = re.compile(r"^[-+]?[\d.,]+%?$")
_ID_REGEX = re.compile(r"^(#?\d+|[A-Z]{2,}-\d+|\d{1,2}/\d{1,2}/\d{2,4})$", re.IGNORECASE)


def detect_column_alignment(values: List[str]) -> WD_ALIGN_PARAGRAPH:
    """
    Intelligently detects whether a column should be centered, right-aligned, or left-aligned:
    - Pure numbers, percentages, currencies -> RIGHT
    - Short IDs, STT, dates -> CENTER
    - Text sentences -> LEFT
    """
    non_empty = [v.strip() for v in values if v and v.strip()]
    if not non_empty:
        return WD_ALIGN_PARAGRAPH.LEFT

    num_count = sum(1 for v in non_empty if _NUMBER_REGEX.match(v) and len(v) <= 15)
    id_count = sum(1 for v in non_empty if _ID_REGEX.match(v) and len(v) <= 12)

    total = len(non_empty)
    if num_count / total >= 0.7:
        return WD_ALIGN_PARAGRAPH.RIGHT
    if id_count / total >= 0.7:
        return WD_ALIGN_PARAGRAPH.CENTER
    return WD_ALIGN_PARAGRAPH.LEFT


def format_table_modern_clean(
    table: Any,
    theme: Optional[DocxTheme] = None,
    col_alignments: Optional[List[WD_ALIGN_PARAGRAPH]] = None,
) -> None:
    """
    Applies the Modern Clean Studio styling to any OpenXML table:
    - cantSplit on all rows, tblHeader on row 0.
    - Ultra-thin borders (#E2E8F0).
    - Roomy cell padding (6pt top/bottom, 8pt left/right).
    - Theme header fill (#0F172A) with white bold text.
    - Zebra striping on alternating data rows (#F8FAFC).
    - Intelligent alignment per column.
    """
    th = theme or THEME_MODERN_TECH
    rows = table.rows
    if not rows:
        return

    # Compute alignments if not explicitly passed
    if col_alignments is None and len(rows) > 1:
        num_cols = len(rows[0].cells)
        col_alignments = []
        for c_idx in range(num_cols):
            col_vals = [r.cells[c_idx].text for r in rows[1:] if c_idx < len(r.cells)]
            col_alignments.append(detect_column_alignment(col_vals))

    pad_dxa_v = str(int(th.table_pad_top_pt * 20))
    pad_dxa_h = str(int(th.table_pad_left_pt * 20))

    for r_idx, row in enumerate(rows):
        trPr = row._tr.get_or_add_trPr()
        if trPr.find(qn("w:cantSplit")) is None:
            trPr.append(parse_xml(f'<w:cantSplit {NS_W}/>'))
        if r_idx == 0 and trPr.find(qn("w:tblHeader")) is None:
            trPr.append(parse_xml(f'<w:tblHeader {NS_W}/>'))

        is_header = (r_idx == 0)
        row_fill = th.table_header_fill if is_header else (th.table_zebra_fill if r_idx % 2 == 1 else "FFFFFF")

        for c_idx, cell in enumerate(row.cells):
            tcPr = cell._tc.get_or_add_tcPr()

            # Sequence: <w:shd> MUST precede <w:vAlign>
            shd = parse_xml(f'<w:shd {NS_W} w:fill="{row_fill}"/>')
            tcPr.append(shd)

            if tcPr.find(qn("w:vAlign")) is None:
                tcPr.append(parse_xml(f'<w:vAlign {NS_W} w:val="center"/>'))

            # Subtle clean borders
            borders = parse_xml(
                f'<w:tcBorders {NS_W}>'
                f'  <w:top w:val="single" w:sz="4" w:space="0" w:color="{th.table_border_color}"/>'
                f'  <w:bottom w:val="single" w:sz="4" w:space="0" w:color="{th.table_border_color}"/>'
                f'  <w:left w:val="single" w:sz="4" w:space="0" w:color="{th.table_border_color}"/>'
                f'  <w:right w:val="single" w:sz="4" w:space="0" w:color="{th.table_border_color}"/>'
                f'</w:tcBorders>'
            )
            tcPr.append(borders)

            # Padding
            tcMar = parse_xml(
                f'<w:tcMar {NS_W}>'
                f'  <w:top w:w="{pad_dxa_v}" w:type="dxa"/>'
                f'  <w:bottom w:w="{pad_dxa_v}" w:type="dxa"/>'
                f'  <w:left w:w="{pad_dxa_h}" w:type="dxa"/>'
                f'  <w:right w:w="{pad_dxa_h}" w:type="dxa"/>'
                f'</w:tcMar>'
            )
            tcPr.append(tcMar)

            # Apply paragraph alignment and font colors
            for p in cell.paragraphs:
                p.paragraph_format.space_before = Pt(1.5)
                p.paragraph_format.space_after = Pt(1.5)
                p.paragraph_format.line_spacing = 1.15

                if is_header:
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    for r in p.runs:
                        r.bold = True
                        r.font.name = th.font_heading
                        r.font.size = Pt(9.5)
                        r.font.color.rgb = th.hex_to_rgb(th.table_header_text)
                else:
                    if col_alignments and c_idx < len(col_alignments):
                        p.alignment = col_alignments[c_idx]
                    for r in p.runs:
                        r.font.name = th.font_primary
                        r.font.size = Pt(9.0)
                        if not r.font.color.rgb:
                            r.font.color.rgb = th.hex_to_rgb(th.color_body)


# ---------------------------------------------------------------------------
# 8. Template Token Formatters (Date, Currency, Number)
# ---------------------------------------------------------------------------

def format_token_value(val: Any, spec: str = "") -> str:
    """
    Formats token value dynamically based on spec:
    - {{DATE:dd/MM/yyyy}}
    - {{NUM:#,##0}} or {{CURRENCY:VND}}
    """
    if val is None:
        return ""

    if not spec:
        return str(val)

    clean_spec = spec.strip().lower()

    # Date formatting
    if clean_spec.startswith("date"):
        try:
            if isinstance(val, (datetime.date, datetime.datetime)):
                d_obj = val
            else:
                d_obj = datetime.datetime.fromisoformat(str(val))
            if "vn" in clean_spec:
                return f"{d_obj.strftime('%d')} tháng {d_obj.strftime('%m')} năm {d_obj.strftime('%Y')}"
            return d_obj.strftime("%d/%m/%Y")
        except Exception:
            return str(val)

    # Currency formatting
    if "vnd" in clean_spec or "currency" in clean_spec or "money" in clean_spec:
        try:
            num = float(re.sub(r"[^\d.-]", "", str(val)))
            return f"{num:,.0f} ₫".replace(",", ".")
        except Exception:
            return str(val)
    if "usd" in clean_spec:
        try:
            num = float(re.sub(r"[^\d.-]", "", str(val)))
            return f"${num:,.2f}"
        except Exception:
            return str(val)

    # Number formatting
    if clean_spec == "int" or clean_spec.startswith("int"):
        try:
            num = float(re.sub(r"[^\d.-]", "", str(val)))
            return f"{round(num):,}"
        except Exception:
            return str(val)

    if clean_spec.startswith("num") or "#" in clean_spec:
        try:
            num = float(re.sub(r"[^\d.-]", "", str(val)))
            return f"{num:,.2f}" if "." in str(val) else f"{num:,.0f}"
        except Exception:
            return str(val)

    return str(val)
