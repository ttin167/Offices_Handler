"""
pptx_writer.py — High-Fidelity OpenXML Presentation Generation Engine (v3.0).

Conforms to Enterprise PowerPoint Invariants:
- PPTX_INV_01: Strict 16:9 Widescreen (13.333 x 7.5 inches)
- PPTX_INV_02: Aspect-Ratio Lock on Technical Diagrams & Images
- PPTX_INV_03: Template-driven Master Slide Inheritance (.potx / .pptx)
- PPTX_INV_04: Zero Text Overflow & Word Wrap Enforcement
- PPTX_INV_05: Mathematical Centering for Technical Architecture Diagrams
- PPTX_INV_06: Native OpenXML Tables with Styled Headers & Zebra Striping
- PPTX_INV_07: Windows Word/PowerPoint File Lock Resilience
- PPTX_INV_08: Presenter Notes Preservation
"""

from __future__ import annotations

import os
import re
import sys
import time
from typing import Any, Dict, List, Optional, Tuple, Union

try:
    from PIL import Image
except ImportError:
    Image = None  # type: ignore

from lxml import etree
from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls

import pptx
from pptx import Presentation
from pptx.presentation import Presentation as PresentationType
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

# ---------------------------------------------------------------------------
# Constants & Color Palettes
# ---------------------------------------------------------------------------

SLIDE_WIDTH_IN = 13.333
SLIDE_HEIGHT_IN = 7.5

THEMES: Dict[str, Dict[str, Any]] = {
    "corporate_blue": {
        "bg_color": RGBColor(248, 250, 252),        # #f8fafc
        "title_color": RGBColor(0, 0, 0),           # Pure Black (PPTX_INV_10)
        "subtitle_color": RGBColor(0, 0, 0),        # Pure Black (PPTX_INV_10)
        "text_color": RGBColor(0, 0, 0),            # Pure Black (PPTX_INV_10)
        "accent_color": RGBColor(37, 99, 235),      # #2563eb Cobalt Blue
        "accent_light": RGBColor(239, 246, 255),    # #eff6ff Blue-50
        "card_bg": RGBColor(255, 255, 255),         # #ffffff
        "card_border": RGBColor(203, 213, 225),     # #cbd5e1
        "font_heading": "Calibri",
        "font_body": "Calibri",
    },
    "modern_dark": {
        "bg_color": RGBColor(15, 23, 42),           # #0f172a Slate-900
        "title_color": RGBColor(56, 189, 248),      # #38bdf8 Cyan Neon
        "subtitle_color": RGBColor(148, 163, 184),  # #94a3b8 Slate-400
        "text_color": RGBColor(241, 245, 249),      # #f1f5f9
        "accent_color": RGBColor(56, 189, 248),     # #38bdf8
        "accent_light": RGBColor(30, 41, 59),       # #1e293b
        "card_bg": RGBColor(30, 41, 59),            # #1e293b
        "card_border": RGBColor(51, 65, 85),        # #334155
        "font_heading": "Arial",
        "font_body": "Arial",
    },
    "academic_light": {
        "bg_color": RGBColor(255, 255, 255),        # #ffffff
        "title_color": RGBColor(0, 0, 0),           # Pure Black (PPTX_INV_10)
        "subtitle_color": RGBColor(0, 0, 0),        # Pure Black (PPTX_INV_10)
        "text_color": RGBColor(0, 0, 0),            # Pure Black (PPTX_INV_10)
        "accent_color": RGBColor(2, 132, 199),      # #0284c7 Sky-600
        "accent_light": RGBColor(240, 249, 255),    # #f0f9ff
        "card_bg": RGBColor(248, 250, 252),         # #f8fafc
        "card_border": RGBColor(226, 232, 240),     # #e2e8f0
        "font_heading": "Times New Roman",
        "font_body": "Times New Roman",
    },
    "thesis_blue": {
        "bg_color": RGBColor(248, 250, 252),        # #f8fafc
        "title_color": RGBColor(0, 0, 0),           # Pure Black (PPTX_INV_10)
        "subtitle_color": RGBColor(0, 0, 0),        # Pure Black (PPTX_INV_10)
        "text_color": RGBColor(0, 0, 0),            # Pure Black (PPTX_INV_10)
        "accent_color": RGBColor(0, 81, 226),       # #0051e2 Royal Blue
        "accent_light": RGBColor(239, 246, 255),    # #eff6ff Blue-50
        "card_bg": RGBColor(255, 255, 255),         # #ffffff
        "card_border": RGBColor(203, 213, 225),     # #cbd5e1
        "font_heading": "Tahoma",
        "font_body": "Calibri",
    },
}


def get_theme(theme_name: str) -> Dict[str, Any]:
    return THEMES.get(theme_name, THEMES["corporate_blue"])


# ---------------------------------------------------------------------------
# OpenXML Sanitization & Text Helpers (PPTX_INV_09 - 13)
# ---------------------------------------------------------------------------

A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"


def set_paragraph_text_clean(
    paragraph: Any,
    text: str,
    font_size_pt: float = 14.0,
    is_bold: bool = False,
    is_italic: bool = False,
    color: Optional[Union[RGBColor, str]] = None,
    font_name: str = "Arial"
) -> Any:
    """
    Populates a paragraph with text runs conforming strictly to PPTX_INV_09 (Run-Level Formatting
    Supremacy) and PPTX_INV_10 (Default Pure Black Typography).
    Supports inline LaTeX math ($...$) via pptx_math.
    """
    # 1. Clean existing runs in paragraph to prevent ghost runs
    for r in list(paragraph._p.findall(f'.//{{{A_NS}}}r')):
        paragraph._p.remove(r)

    # 2. Strip defRPr bold/color locks from paragraph
    pPr = paragraph._p.find(f'{{{A_NS}}}pPr')
    if pPr is not None:
        defRPr = pPr.find(f'{{{A_NS}}}defRPr')
        if defRPr is not None:
            if 'b' in defRPr.attrib:
                del defRPr.attrib['b']
            if 'i' in defRPr.attrib:
                del defRPr.attrib['i']
            for sf in defRPr.findall(f'{{{A_NS}}}solidFill'):
                defRPr.remove(sf)

    hex_clr = "000000"
    if isinstance(color, RGBColor):
        hex_clr = f"{color[0]:02X}{color[1]:02X}{color[2]:02X}"
    elif isinstance(color, str):
        hex_clr = color.lstrip("#")

    if "$" in text:
        try:
            from .pptx_math import add_math_text_to_paragraph
        except (ImportError, ValueError):
            from pptx_math import add_math_text_to_paragraph  # type: ignore
        add_math_text_to_paragraph(
            paragraph,
            text,
            base_sz=int(round(font_size_pt * 100)),
            is_bold=is_bold,
            color=hex_clr,
            typeface=font_name
        )
        return paragraph
    else:
        run = paragraph.add_run()
        run.text = text
        run.font.name = font_name
        run.font.size = Pt(font_size_pt)
        run.font.bold = is_bold
        run.font.italic = is_italic
        if isinstance(color, RGBColor):
            run.font.color.rgb = color
        elif isinstance(color, str):
            h = color.lstrip("#")
            run.font.color.rgb = RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))
        else:
            run.font.color.rgb = RGBColor(0, 0, 0)
        return run


def sanitize_openxml_presentation(prs: Any) -> int:
    """
    Sanitizes an entire PowerPoint presentation against DrawingML / OpenXML defects.
    Conforms to:
    - PPTX_INV_09: Run-Level Formatting Supremacy (strips b="1", b="0", solidFill from defRPr)
    - PPTX_INV_10: Default Pure Black Typography (replaces legacy palette #0F294A, #475569 with #000000)
    - PPTX_INV_11: Slide Numbering Freedom (sets clean run-level black 12pt Arial formatting)
    - PPTX_INV_13: AlternateContent Ghost Shape Pruning (removes orphan AlternateContent blocks)
    """
    sanitized_count = 0
    for slide in prs.slides:
        # 1. Prune ghost AlternateContent in spTree (PPTX_INV_13)
        sp_tree = slide._element.spTree
        for child in list(sp_tree):
            if child.tag.endswith('AlternateContent'):
                xml_str = etree.tostring(child).decode('utf-8', errors='ignore')
                if '0F294A' in xml_str.upper() or '475569' in xml_str.upper() or 'PTC' in xml_str:
                    sp_tree.remove(child)
                    sanitized_count += 1

        # 2. Inspect all shapes for slide numbers (PPTX_INV_11)
        for shape in slide.shapes:
            if 'slide number' in shape.name.lower() or 'date placeholder' in shape.name.lower() or (
                shape.has_text_frame and shape.text.strip().isdigit() and len(shape.text.strip()) <= 3 and shape.name.startswith('Slide Number')
            ) or (shape.is_placeholder and hasattr(shape, 'placeholder_format') and shape.placeholder_format.type == 12):
                for p in shape._element.findall(f'.//{{{A_NS}}}p'):
                    pPr = p.find(f'{{{A_NS}}}pPr')
                    if pPr is not None:
                        defRPr = pPr.find(f'{{{A_NS}}}defRPr')
                        if defRPr is not None:
                            for sf in defRPr.findall(f'{{{A_NS}}}solidFill'):
                                defRPr.remove(sf)
                                sanitized_count += 1
                            if 'b' in defRPr.attrib:
                                del defRPr.attrib['b']
                                sanitized_count += 1
                            if 'i' in defRPr.attrib:
                                del defRPr.attrib['i']
                    for r in p.findall(f'{{{A_NS}}}r'):
                        rPr = r.find(f'{{{A_NS}}}rPr')
                        if rPr is None:
                            rPr = parse_xml(f'<a:rPr {nsdecls("a")}/>')
                            r.insert(0, rPr)
                        rPr.set('sz', '1200')
                        rPr.set('b', '0')
                        rPr.set('i', '0')
                        for sf in rPr.findall(f'{{{A_NS}}}solidFill'):
                            rPr.remove(sf)
                        rPr.append(parse_xml(f'<a:solidFill {nsdecls("a")}><a:srgbClr val="000000"/></a:solidFill>'))
                        for lat in rPr.findall(f'{{{A_NS}}}latin'):
                            rPr.remove(lat)
                        rPr.append(parse_xml(f'<a:latin {nsdecls("a")} typeface="Arial"/>'))
                        sanitized_count += 1

        # 3. Universal Slide Paragraph Sanitization (PPTX_INV_09 & PPTX_INV_10)
        for p in slide._element.findall(f'.//{{{A_NS}}}p'):
            pPr = p.find(f'{{{A_NS}}}pPr')
            if pPr is not None:
                defRPr = pPr.find(f'{{{A_NS}}}defRPr')
                if defRPr is not None:
                    if 'b' in defRPr.attrib:
                        del defRPr.attrib['b']
                        sanitized_count += 1
                    if 'i' in defRPr.attrib:
                        del defRPr.attrib['i']
                    for sf in defRPr.findall(f'{{{A_NS}}}solidFill'):
                        for clr in sf.findall(f'.//{{{A_NS}}}srgbClr'):
                            if clr.attrib.get('val', '').upper() in ['0F294A', '475569', '64748B']:
                                defRPr.remove(sf)
                                sanitized_count += 1
                                break

            # Clean runs: replace legacy hardcoded navy/slate colors in rPr with pure black
            for r in p.findall(f'{{{A_NS}}}r'):
                rPr = r.find(f'{{{A_NS}}}rPr')
                if rPr is not None:
                    for clr in rPr.findall(f'.//{{{A_NS}}}srgbClr'):
                        if clr.attrib.get('val', '').upper() in ['0F294A', '475569']:
                            clr.attrib['val'] = '000000'
                            sanitized_count += 1

    return sanitized_count


# ---------------------------------------------------------------------------
# Resilience & File Lock Management
# ---------------------------------------------------------------------------

def safe_save_pptx(prs: PresentationType, target_path: str) -> str:
    """
    Saves a PowerPoint presentation. If the target file is locked by MS PowerPoint,
    gracefully saves to a timestamped or fallback path to prevent crash (PPTX_INV_07).
    Automatically runs OpenXML pre-save sanitization to enforce PPTX_INV_09, 10, 11, 13.
    """
    # Enforce OpenXML Invariants before saving
    sanitize_openxml_presentation(prs)

    target_path = os.path.abspath(target_path)
    os.makedirs(os.path.dirname(target_path), exist_ok=True)

    try:
        prs.save(target_path)
        return target_path
    except PermissionError:
        base, ext = os.path.splitext(target_path)
        ts = int(time.time())
        fallback_path = f"{base}_updated_{ts}{ext}"
        print(f"[WARN] File is locked by PowerPoint: {target_path}", file=sys.stderr)
        print(f"[INFO] Saving safely to fallback: {fallback_path}", file=sys.stderr)
        prs.save(fallback_path)
        return fallback_path


# ---------------------------------------------------------------------------
# Presentation Initialization
# ---------------------------------------------------------------------------

def init_presentation(template_path: Optional[str] = None) -> PresentationType:
    """
    Initializes a PowerPoint presentation. Enforces 16:9 widescreen format (PPTX_INV_01).
    Inherits Slide Master if template_path is provided (PPTX_INV_03).
    """
    if template_path and os.path.isfile(template_path):
        prs = Presentation(template_path)
    else:
        prs = Presentation()

    prs.slide_width = Inches(SLIDE_WIDTH_IN)
    prs.slide_height = Inches(SLIDE_HEIGHT_IN)
    return prs


def _apply_slide_background(slide: Any, color: RGBColor) -> None:
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = color


def _calculate_image_fit(
    image_path: str,
    box_left: float,
    box_top: float,
    box_width: float,
    box_height: float,
) -> Tuple[float, float, float, float]:
    """Calculates aspect-ratio locked (x, y, w, h) in inches mathematically centered inside bounding box."""
    orig_w, orig_h = 1000, 600
    if Image is not None and os.path.isfile(image_path):
        try:
            with Image.open(image_path) as im:
                orig_w, orig_h = im.size
        except Exception:
            pass

    if orig_h <= 0:
        orig_h = 600

    aspect = orig_w / float(orig_h)
    box_aspect = box_width / max(0.1, box_height)

    if aspect >= box_aspect:
        final_w = box_width
        final_h = box_width / aspect
    else:
        final_h = box_height
        final_w = box_height * aspect

    final_x = box_left + (box_width - final_w) / 2.0
    final_y = box_top + (box_height - final_h) / 2.0
    return final_x, final_y, final_w, final_h


def _format_bullet_paragraph(
    p: Any,
    item: Union[str, Tuple[Any, ...], Dict[str, Any]],
    theme: Dict[str, Any],
    font_size_pt: float = 15.5,
    prefix_color: Optional[RGBColor] = None,
    bullet_color: Optional[RGBColor] = None,
    body_color: Optional[RGBColor] = None,
) -> None:
    """Formats bullet paragraph with high-contrast prefix runs, inline math support, and zero text overflow."""
    p.space_after = Pt(max(3.0, font_size_pt * 0.35))
    p.line_spacing = 1.15
    p.alignment = PP_ALIGN.LEFT

    # Strip any defRPr locks
    pPr = p._p.find(f'{{{A_NS}}}pPr')
    if pPr is not None:
        defRPr = pPr.find(f'{{{A_NS}}}defRPr')
        if defRPr is not None:
            if 'b' in defRPr.attrib:
                del defRPr.attrib['b']
            if 'i' in defRPr.attrib:
                del defRPr.attrib['i']
            for sf in defRPr.findall(f'{{{A_NS}}}solidFill'):
                defRPr.remove(sf)

    prefix = ""
    body = ""
    level = 0

    if isinstance(item, (tuple, list)):
        if len(item) == 2:
            if isinstance(item[1], int):
                body, level = str(item[0]), item[1]
            else:
                prefix, body = str(item[0]), str(item[1])
        elif len(item) >= 3:
            prefix, body, level = str(item[0]), str(item[1]), int(item[2])
    elif isinstance(item, dict):
        prefix = str(item.get("prefix", item.get("title", "")))
        body = str(item.get("text", item.get("body", item.get("desc", ""))))
        level = int(item.get("level", 0))
    elif isinstance(item, str):
        # Match markdown bold prefix like "**Prefix:** Body" or "**Prefix**: Body"
        m = re.match(r'^\*\*(.*?)\*\*[:\s]*(.*)$', item.strip())
        if m:
            prefix = m.group(1).strip()
            body = m.group(2).strip()
        elif ":" in item and len(item.split(":", 1)[0]) <= 32 and not item.startswith("http"):
            parts = item.split(":", 1)
            prefix = parts[0].strip()
            body = parts[1].strip()
        else:
            body = item.strip()

    p.level = min(level, 2)
    eff_size = font_size_pt if p.level == 0 else max(11.0, font_size_pt - 2.0)

    # Bullet dot run
    r_dot = p.add_run()
    r_dot.text = "•  " if p.level == 0 else "   –  "
    r_dot.font.name = theme.get("font_body", "Calibri")
    r_dot.font.size = Pt(eff_size)
    r_dot.font.bold = True
    r_dot.font.color.rgb = bullet_color or theme.get("accent_color", RGBColor(37, 99, 235))

    # Prefix run (bold)
    if prefix:
        clean_pre = prefix.rstrip(":") + ": "
        if "$" in clean_pre:
            try:
                from .pptx_math import add_math_text_to_paragraph
            except (ImportError, ValueError):
                from pptx_math import add_math_text_to_paragraph  # type: ignore
            add_math_text_to_paragraph(
                p,
                clean_pre,
                base_sz=int(round(eff_size * 100)),
                is_bold=True,
                color="000000" if prefix_color is None else (
                    f"{prefix_color[0]:02X}{prefix_color[1]:02X}{prefix_color[2]:02X}"
                    if isinstance(prefix_color, RGBColor) else str(prefix_color).lstrip("#")
                ),
                typeface=theme.get("font_body", "Calibri")
            )
        else:
            r_pre = p.add_run()
            r_pre.text = clean_pre
            r_pre.font.name = theme.get("font_body", "Calibri")
            r_pre.font.size = Pt(eff_size)
            r_pre.font.bold = True
            r_pre.font.color.rgb = prefix_color or theme.get("title_color", RGBColor(0, 0, 0))

    # Body run (regular, with math parsing if $ present)
    if "$" in body:
        try:
            from .pptx_math import add_math_text_to_paragraph
        except (ImportError, ValueError):
            from pptx_math import add_math_text_to_paragraph  # type: ignore
        add_math_text_to_paragraph(
            p,
            body,
            base_sz=int(round(eff_size * 100)),
            is_bold=False,
            color="000000" if body_color is None else (
                f"{body_color[0]:02X}{body_color[1]:02X}{body_color[2]:02X}"
                if isinstance(body_color, RGBColor) else str(body_color).lstrip("#")
            ),
            typeface=theme.get("font_body", "Calibri")
        )
    else:
        r_body = p.add_run()
        r_body.text = body
        r_body.font.name = theme.get("font_body", "Calibri")
        r_body.font.size = Pt(eff_size)
        r_body.font.bold = False
        r_body.font.color.rgb = body_color or theme.get("text_color", RGBColor(0, 0, 0))


def _clean_slide_content(slide: Any) -> None:
    """Removes existing content shapes from a template slide while preserving master branding and placeholders."""
    to_delete = []
    for sh in slide.shapes:
        if sh.name == 'Rectangle 1' and sh.top < 100000:
            continue
        if sh.name == 'Picture 2' and sh.left == 0:
            continue
        if sh.is_placeholder and hasattr(sh, 'placeholder_format') and sh.placeholder_format.type == 12:
            continue
        if 'Slide Number' in sh.name:
            continue
        to_delete.append(sh)
    for sh in to_delete:
        sp = sh._element
        parent = sp.getparent()
        if parent is not None:
            parent.remove(sp)


def _update_slide_number(slide: Any, num: int, theme: Dict[str, Any]) -> None:
    """Updates slide number placeholder conforming strictly to PPTX_INV_11 (Slide Numbering Freedom)."""
    for sh in slide.shapes:
        if 'Slide Number' in sh.name or (sh.is_placeholder and hasattr(sh, 'placeholder_format') and sh.placeholder_format.type == 12):
            tf = sh.text_frame
            tf.clear()
            p = tf.paragraphs[0]
            # Strip defRPr lock
            pPr = p._p.find(f'{{{A_NS}}}pPr')
            if pPr is not None:
                defRPr = pPr.find(f'{{{A_NS}}}defRPr')
                if defRPr is not None:
                    pPr.remove(defRPr)
            run = p.add_run()
            run.text = str(num)
            run.font.name = "Arial"
            run.font.size = Pt(12)
            run.font.bold = False
            run.font.italic = False
            run.font.color.rgb = RGBColor(0, 0, 0)


def _prepare_slide(
    prs: PresentationType,
    theme: Dict[str, Any],
    target_slide: Optional[Any] = None,
) -> Any:
    """Prepares a slide either by cleaning an existing target_slide or adding a new blank slide."""
    if target_slide is not None:
        _clean_slide_content(target_slide)
        return target_slide
    blank_layout = prs.slide_layouts[6] if len(prs.slide_layouts) > 6 else prs.slide_layouts[0]
    slide = prs.slides.add_slide(blank_layout)
    _apply_slide_background(slide, theme["bg_color"])
    return slide


def _add_slide_header(
    slide: Any,
    title: str,
    subtitle: Optional[str],
    theme: Dict[str, Any],
    banner: Optional[str] = None,
) -> float:
    """Adds a standard or 2-tier top banner/title block to a slide. Returns top_y in inches for body content."""
    # Detect if slide has template header branding (logo or top bar)
    has_branding = any(
        (sh.name == 'Rectangle 1' and sh.top < 100000) or (sh.name == 'Picture 2' and sh.left == 0)
        for sh in slide.shapes
    )

    if banner:
        if has_branding:
            # Banner positioned cleanly to the right of logo
            tb_banner = slide.shapes.add_textbox(Inches(2.0), Inches(0.12), Inches(9.5), Inches(0.6))
            tf_b = tb_banner.text_frame
            tf_b.word_wrap = True
            tf_b.margin_left = Inches(0)
            tf_b.margin_top = Inches(0)
            p_b = tf_b.paragraphs[0]
            set_paragraph_text_clean(
                p_b, banner, font_size_pt=24, is_bold=True,
                color=theme.get("accent_color", RGBColor(0, 81, 226)),
                font_name=theme.get("font_heading", "Tahoma")
            )

            # Title positioned below banner with exact 28pt bold
            tb_title = slide.shapes.add_textbox(Inches(0.8), Inches(0.89), Inches(11.85), Inches(0.65))
            tf_t = tb_title.text_frame
            tf_t.word_wrap = True
            tf_t.margin_left = Inches(0)
            tf_t.margin_top = Inches(0)
            p_t = tf_t.paragraphs[0]
            set_paragraph_text_clean(
                p_t, title, font_size_pt=28, is_bold=True,
                color=theme.get("title_color", RGBColor(0, 0, 0)),
                font_name=theme.get("font_heading", "Tahoma")
            )

            if subtitle:
                p_sub = tf_t.add_paragraph()
                set_paragraph_text_clean(
                    p_sub, subtitle, font_size_pt=13, is_bold=False,
                    color=theme.get("subtitle_color", RGBColor(0, 0, 0)),
                    font_name=theme.get("font_body", "Calibri")
                )
                return 1.85
            return 1.66
        else:
            # Standard canvas 2-tier banner
            tb_banner = slide.shapes.add_textbox(Inches(0.8), Inches(0.2), Inches(11.733), Inches(0.55))
            tf_b = tb_banner.text_frame
            tf_b.word_wrap = True
            tf_b.margin_left = Inches(0)
            tf_b.margin_top = Inches(0)
            p_b = tf_b.paragraphs[0]
            p_b.alignment = PP_ALIGN.CENTER
            set_paragraph_text_clean(
                p_b, banner, font_size_pt=20, is_bold=True,
                color=theme.get("accent_color", RGBColor(0, 81, 226)),
                font_name=theme.get("font_heading", "Tahoma")
            )

            tb_sub = slide.shapes.add_textbox(Inches(0.8), Inches(0.8), Inches(11.733), Inches(0.7))
            tf_s = tb_sub.text_frame
            tf_s.word_wrap = True
            tf_s.margin_left = Inches(0)
            tf_s.margin_top = Inches(0)
            p_s = tf_s.paragraphs[0]
            set_paragraph_text_clean(
                p_s, title, font_size_pt=28, is_bold=True,
                color=theme.get("title_color", RGBColor(0, 0, 0)),
                font_name=theme.get("font_heading", "Tahoma")
            )

            if subtitle:
                p2 = tf_s.add_paragraph()
                p2.space_before = Pt(3)
                set_paragraph_text_clean(
                    p2, subtitle, font_size_pt=13, is_bold=False,
                    color=theme.get("subtitle_color", RGBColor(0, 0, 0)),
                    font_name=theme.get("font_body", "Calibri")
                )
                return 1.85
            return 1.66
    else:
        tx_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.5), Inches(11.7), Inches(1.2))
        tf = tx_box.text_frame
        tf.word_wrap = True
        tf.margin_left = Inches(0)
        tf.margin_top = Inches(0)

        p = tf.paragraphs[0]
        set_paragraph_text_clean(
            p, title, font_size_pt=28, is_bold=True,
            color=theme.get("title_color", RGBColor(0, 0, 0)),
            font_name=theme.get("font_heading", "Tahoma")
        )

        if subtitle:
            p2 = tf.add_paragraph()
            p2.space_before = Pt(4)
            set_paragraph_text_clean(
                p2, subtitle, font_size_pt=13, is_bold=False,
                color=theme.get("subtitle_color", RGBColor(0, 0, 0)),
                font_name=theme.get("font_body", "Calibri")
            )
            return 1.9
        return 1.7


# ---------------------------------------------------------------------------
# Slide Builders
# ---------------------------------------------------------------------------

def add_title_slide(
    prs: PresentationType,
    title: str,
    subtitle: Optional[str] = None,
    presenter: Optional[str] = None,
    date_str: Optional[str] = None,
    theme_name: str = "corporate_blue",
    target_slide: Optional[Any] = None,
) -> Any:
    """Creates a premium Widescreen Title Slide."""
    theme = get_theme(theme_name)
    slide = _prepare_slide(prs, theme, target_slide=target_slide)

    # Decorative Accent Bar
    accent_bar = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        Inches(1.0), Inches(2.2), Inches(0.15), Inches(3.0)
    )
    accent_bar.fill.solid()
    accent_bar.fill.fore_color.rgb = theme["accent_color"]
    accent_bar.line.color.rgb = theme["accent_color"]

    # Main text container
    tx_box = slide.shapes.add_textbox(Inches(1.4), Inches(2.0), Inches(10.5), Inches(3.5))
    tf = tx_box.text_frame
    tf.word_wrap = True

    # Title
    p_title = tf.paragraphs[0]
    set_paragraph_text_clean(
        p_title, title, font_size_pt=40, is_bold=True,
        color=theme.get("title_color", RGBColor(0, 0, 0)),
        font_name=theme["font_heading"]
    )

    # Subtitle
    if subtitle:
        p_sub = tf.add_paragraph()
        p_sub.space_before = Pt(14)
        set_paragraph_text_clean(
            p_sub, subtitle, font_size_pt=20, is_bold=False,
            color=theme.get("subtitle_color", RGBColor(0, 0, 0)),
            font_name=theme["font_body"]
        )

    # Presenter & Meta info
    meta_parts = []
    if presenter:
        meta_parts.append(presenter)
    if date_str:
        meta_parts.append(date_str)

    if meta_parts:
        p_meta = tf.add_paragraph()
        p_meta.space_before = Pt(28)
        set_paragraph_text_clean(
            p_meta, "  |  ".join(meta_parts), font_size_pt=14, is_bold=True,
            color=theme.get("accent_color", RGBColor(0, 81, 226)),
            font_name=theme["font_body"]
        )

    return slide


def add_chapter_slide(
    prs: PresentationType,
    chapter_num: str,
    title: str,
    subtitle: Optional[str] = None,
    theme_name: str = "thesis_blue",
    target_slide: Optional[Any] = None,
) -> Any:
    """Creates a prominent Chapter Transition / Divider slide."""
    theme = get_theme(theme_name)
    slide = _prepare_slide(prs, theme, target_slide=target_slide)

    # Decorative Accent Bar
    accent_bar = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        Inches(1.2), Inches(2.2), Inches(0.18), Inches(2.8)
    )
    accent_bar.fill.solid()
    accent_bar.fill.fore_color.rgb = theme["accent_color"]
    accent_bar.line.color.rgb = theme["accent_color"]

    # Main text box
    tx_box = slide.shapes.add_textbox(Inches(1.7), Inches(2.0), Inches(10.2), Inches(3.2))
    tf = tx_box.text_frame
    tf.word_wrap = True

    # Chapter Number / Category
    p_num = tf.paragraphs[0]
    set_paragraph_text_clean(
        p_num, chapter_num.upper(), font_size_pt=18, is_bold=True,
        color=theme.get("accent_color", RGBColor(0, 81, 226)),
        font_name=theme["font_heading"]
    )

    # Chapter Title
    p_title = tf.add_paragraph()
    p_title.space_before = Pt(10)
    set_paragraph_text_clean(
        p_title, title, font_size_pt=36, is_bold=True,
        color=theme.get("title_color", RGBColor(0, 0, 0)),
        font_name=theme["font_heading"]
    )

    # Subtitle
    if subtitle:
        p_sub = tf.add_paragraph()
        p_sub.space_before = Pt(12)
        set_paragraph_text_clean(
            p_sub, subtitle, font_size_pt=16, is_bold=False,
            color=theme.get("subtitle_color", RGBColor(0, 0, 0)),
            font_name=theme["font_body"]
        )

    return slide


def add_content_slide(
    prs: PresentationType,
    title: str,
    bullets: List[Any],
    subtitle: Optional[str] = None,
    banner: Optional[str] = None,
    theme_name: str = "corporate_blue",
    bullet_font_size: float = 17.0,
    use_card: bool = True,
    target_slide: Optional[Any] = None,
) -> Any:
    """Creates a Content Slide with hierarchical multi-level bullets in a styled card."""
    theme = get_theme(theme_name)
    slide = _prepare_slide(prs, theme, target_slide=target_slide)
    top_y = _add_slide_header(slide, title, subtitle, theme, banner=banner)

    content_top = top_y + 0.08
    available_h = max(3.5, 7.0 - content_top)

    if use_card:
        card = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE,
            Inches(0.8), Inches(content_top), Inches(11.733), Inches(available_h)
        )
        card.fill.solid()
        card.fill.fore_color.rgb = theme["card_bg"]
        card.line.color.rgb = theme["card_border"]
        card.line.width = Pt(1.2)

        tf = card.text_frame
        tf.word_wrap = True
        tf.vertical_anchor = MSO_ANCHOR.TOP
        tf.margin_left = Inches(0.3)
        tf.margin_right = Inches(0.3)
        tf.margin_top = Inches(0.3)
        tf.margin_bottom = Inches(0.2)
        tf.clear()
    else:
        tx_box = slide.shapes.add_textbox(Inches(0.8), Inches(content_top), Inches(11.733), Inches(available_h))
        tf = tx_box.text_frame
        tf.word_wrap = True

    # Dynamic 2-Tier font sizing calculation if bullet_font_size is default (16.5 or 17.0)
    effective_font_size = bullet_font_size
    if bullet_font_size in (16.5, 17.0) and bullets:
        b_count = len(bullets)
        total_chars = sum(
            len(str(b.get("text", b) if isinstance(b, dict) else (b[1] if isinstance(b, (tuple, list)) and len(b) > 1 else b)))
            for b in bullets
        )
        if b_count <= 3 and total_chars < 220:
            effective_font_size = 18.0
        elif b_count <= 5 and total_chars < 450:
            effective_font_size = 16.5
        elif b_count <= 7:
            effective_font_size = 15.0
        else:
            effective_font_size = 13.5

    first = True
    for item in bullets:
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        _format_bullet_paragraph(p, item, theme, font_size_pt=effective_font_size)
        if len(bullets) >= 6:
            p.space_after = Pt(6)

    return slide


def add_two_column_slide(
    prs: PresentationType,
    title: str,
    left_bullets: List[Any],
    right_bullets: List[Any],
    subtitle: Optional[str] = None,
    banner: Optional[str] = None,
    theme_name: str = "corporate_blue",
    target_slide: Optional[Any] = None,
) -> Any:
    """Creates a Two-Column Content Slide with separated card panels."""
    theme = get_theme(theme_name)
    slide = _prepare_slide(prs, theme, target_slide=target_slide)
    top_y = _add_slide_header(slide, title, subtitle, theme, banner=banner)

    col_w = Inches(5.6)
    col_h = Inches(max(3.5, 7.0 - top_y - 0.2))

    # Left & Right Columns
    for idx, (col_x, items) in enumerate([
        (Inches(0.8), left_bullets),
        (Inches(6.8), right_bullets),
    ]):
        card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, col_x, Inches(top_y), col_w, col_h)
        card.fill.solid()
        card.fill.fore_color.rgb = theme["card_bg"]
        card.line.color.rgb = theme["card_border"]
        card.line.width = Pt(1)

        tf = card.text_frame
        tf.word_wrap = True
        tf.margin_left = Inches(0.4)
        tf.margin_right = Inches(0.4)
        tf.margin_top = Inches(0.4)

        first = True
        for itm in items:
            p = tf.paragraphs[0] if first else tf.add_paragraph()
            first = False
            _format_bullet_paragraph(p, itm, theme, font_size_pt=15.0)

    return slide


def add_diagram_slide(
    prs: PresentationType,
    title: str,
    image_path: str,
    caption: Optional[str] = None,
    subtitle: Optional[str] = None,
    banner: Optional[str] = None,
    theme_name: str = "corporate_blue",
    target_slide: Optional[Any] = None,
) -> Any:
    """
    Creates a Technical Diagram presentation slide.
    Enforces Aspect-Ratio Locking and Mathematical Centering (PPTX_INV_02, PPTX_INV_05).
    """
    theme = get_theme(theme_name)
    slide = _prepare_slide(prs, theme, target_slide=target_slide)
    top_y = _add_slide_header(slide, title, subtitle, theme, banner=banner)

    if not os.path.isfile(image_path):
        err_box = slide.shapes.add_textbox(Inches(1.0), Inches(top_y + 0.5), Inches(11.3), Inches(3.0))
        p = err_box.text_frame.paragraphs[0]
        p.text = f"[Diagram Image Not Found: {image_path}]"
        p.font.color.rgb = RGBColor(220, 38, 38)
        p.font.bold = True
        return slide

    # Bounding Box Limits (Inches)
    box_left = 0.8
    box_top = top_y + 0.1
    box_width = 11.733
    box_height = 4.8 if caption else 5.1
    box_height = min(box_height, 7.1 - box_top)

    # Aspect Ratio Calculation
    orig_w, orig_h = 1000, 600
    if Image is not None:
        try:
            with Image.open(image_path) as im:
                orig_w, orig_h = im.size
        except Exception:
            pass

    aspect = orig_w / float(orig_h)
    box_aspect = box_width / box_height

    if aspect >= box_aspect:
        final_w = box_width
        final_h = box_width / aspect
    else:
        final_h = box_height
        final_w = box_height * aspect

    final_x = box_left + (box_width - final_w) / 2.0
    final_y = box_top + (box_height - final_h) / 2.0

    slide.shapes.add_picture(
        image_path,
        Inches(final_x),
        Inches(final_y),
        width=Inches(final_w),
        height=Inches(final_h),
    )

    if caption:
        cap_box = slide.shapes.add_textbox(
            Inches(0.8),
            Inches(min(6.9, final_y + final_h + 0.08)),
            Inches(11.733),
            Inches(0.4),
        )
        tf = cap_box.text_frame
        tf.word_wrap = True
        p_cap = tf.paragraphs[0]
        p_cap.alignment = PP_ALIGN.CENTER
        set_paragraph_text_clean(
            p_cap, caption, font_size_pt=11.5, is_italic=True,
            color=theme.get("subtitle_color", RGBColor(0, 0, 0)),
            font_name=theme["font_body"]
        )

    return slide


def add_split_diagram_slide(
    prs: PresentationType,
    title: str,
    bullets: List[Any],
    image_path: str,
    caption: Optional[str] = None,
    subtitle: Optional[str] = None,
    banner: Optional[str] = None,
    theme_name: str = "corporate_blue",
    card_width_ratio: float = 0.50,
    bullet_font_size: float = 15.5,
    target_slide: Optional[Any] = None,
) -> Any:
    """
    Creates a Split-Screen Technical Slide:
    - Left side: High-contrast Card with bold semantic prefix bullets.
    - Right side: Aspect-ratio locked Technical Diagram with centered caption underneath.
    """
    theme = get_theme(theme_name)
    slide = _prepare_slide(prs, theme, target_slide=target_slide)
    top_y = _add_slide_header(slide, title, subtitle, theme, banner=banner)

    # Geometry Layout
    left_x = 0.8
    content_top = top_y + 0.08
    available_w = 11.733
    available_h = max(3.5, 7.0 - content_top)
    gap_x = 0.35

    ratio = max(0.3, min(card_width_ratio, 0.7))
    card_w = (available_w - gap_x) * ratio
    panel_w = available_w - gap_x - card_w
    panel_x = left_x + card_w + gap_x

    # 1. Left Card Container
    card = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(left_x), Inches(content_top), Inches(card_w), Inches(available_h)
    )
    card.fill.solid()
    card.fill.fore_color.rgb = theme["card_bg"]
    card.line.color.rgb = theme["card_border"]
    card.line.width = Pt(1.2)

    tf = card.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.TOP
    tf.margin_left = Inches(0.28)
    tf.margin_right = Inches(0.28)
    tf.margin_top = Inches(0.28)
    tf.margin_bottom = Inches(0.20)
    tf.clear()

    first = True
    for itm in bullets:
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        _format_bullet_paragraph(p, itm, theme, font_size_pt=bullet_font_size)

    # 2. Right Diagram Panel
    if not os.path.isfile(image_path):
        err_box = slide.shapes.add_textbox(Inches(panel_x), Inches(content_top + 0.5), Inches(panel_w), Inches(2.0))
        p = err_box.text_frame.paragraphs[0]
        p.text = f"[Image Not Found: {os.path.basename(image_path)}]"
        p.font.color.rgb = RGBColor(220, 38, 38)
        p.font.bold = True
        return slide

    cap_h = 0.38 if caption else 0.0
    img_box_h = max(1.0, available_h - cap_h)

    ix, iy, iw, ih = _calculate_image_fit(image_path, panel_x, content_top, panel_w, img_box_h)
    slide.shapes.add_picture(image_path, Inches(ix), Inches(iy), width=Inches(iw), height=Inches(ih))

    if caption:
        cap_box = slide.shapes.add_textbox(
            Inches(panel_x), Inches(iy + ih + 0.05), Inches(panel_w), Inches(0.35)
        )
        tf_c = cap_box.text_frame
        tf_c.word_wrap = True
        p_c = tf_c.paragraphs[0]
        p_c.alignment = PP_ALIGN.CENTER
        set_paragraph_text_clean(
            p_c, caption, font_size_pt=11.5, is_italic=True,
            color=theme.get("subtitle_color", RGBColor(0, 0, 0)),
            font_name=theme["font_body"]
        )

    return slide


def add_dual_diagram_slide(
    prs: PresentationType,
    title: str,
    bullets: List[Any],
    img1_path: str,
    cap1: str,
    img2_path: str,
    cap2: str,
    subtitle: Optional[str] = None,
    banner: Optional[str] = None,
    orientation: str = "vertical",
    theme_name: str = "corporate_blue",
    card_width_ratio: float = 0.50,
    bullet_font_size: float = 15.5,
    target_slide: Optional[Any] = None,
) -> Any:
    """
    Creates a Dual-Diagram Comparative Technical Slide:
    - Left side: High-contrast Card with bullets.
    - Right side: Two diagrams (stacked vertically or side-by-side) with individual aspect locks and captions.
    """
    theme = get_theme(theme_name)
    slide = _prepare_slide(prs, theme, target_slide=target_slide)
    top_y = _add_slide_header(slide, title, subtitle, theme, banner=banner)

    # Geometry Layout
    left_x = 0.8
    content_top = top_y + 0.08
    available_w = 11.733
    available_h = max(3.5, 7.0 - content_top)
    gap_x = 0.35

    ratio = max(0.3, min(card_width_ratio, 0.7))
    card_w = (available_w - gap_x) * ratio
    panel_w = available_w - gap_x - card_w
    panel_x = left_x + card_w + gap_x

    # 1. Left Card Container
    card = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(left_x), Inches(content_top), Inches(card_w), Inches(available_h)
    )
    card.fill.solid()
    card.fill.fore_color.rgb = theme["card_bg"]
    card.line.color.rgb = theme["card_border"]
    card.line.width = Pt(1.2)

    tf = card.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.TOP
    tf.margin_left = Inches(0.28)
    tf.margin_right = Inches(0.28)
    tf.margin_top = Inches(0.28)
    tf.margin_bottom = Inches(0.20)
    tf.clear()

    first = True
    for itm in bullets:
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        _format_bullet_paragraph(p, itm, theme, font_size_pt=bullet_font_size)

    # 2. Right Dual-Diagram Panel
    items = [(img1_path, cap1), (img2_path, cap2)]

    if orientation == "horizontal":
        sub_gap = 0.20
        sub_w = (panel_w - sub_gap) / 2.0
        for i, (path, cap) in enumerate(items):
            sub_x = panel_x + i * (sub_w + sub_gap)
            cap_h = 0.38 if cap else 0.0
            avail_h = max(1.0, available_h - cap_h)
            if os.path.isfile(path):
                ix, iy, iw, ih = _calculate_image_fit(path, sub_x, content_top, sub_w, avail_h)
                slide.shapes.add_picture(path, Inches(ix), Inches(iy), width=Inches(iw), height=Inches(ih))
                if cap:
                    cb = slide.shapes.add_textbox(Inches(sub_x), Inches(iy + ih + 0.04), Inches(sub_w), Inches(0.35))
                    tf_c = cb.text_frame
                    tf_c.word_wrap = True
                    p_c = tf_c.paragraphs[0]
                    p_c.alignment = PP_ALIGN.CENTER
                    set_paragraph_text_clean(
                        p_c, cap, font_size_pt=11.0, is_italic=True,
                        color=theme.get("subtitle_color", RGBColor(0, 0, 0)),
                        font_name=theme["font_body"]
                    )
    else:
        # Vertical stacking
        sub_gap = 0.30
        sub_h = (available_h - sub_gap) / 2.0
        for i, (path, cap) in enumerate(items):
            sub_y = content_top + i * (sub_h + sub_gap)
            cap_h = 0.32 if cap else 0.0
            avail_img_h = max(0.8, sub_h - cap_h - 0.05)
            if os.path.isfile(path):
                ix, iy, iw, ih = _calculate_image_fit(path, panel_x, sub_y, panel_w, avail_img_h)
                slide.shapes.add_picture(path, Inches(ix), Inches(iy), width=Inches(iw), height=Inches(ih))
                if cap:
                    cb = slide.shapes.add_textbox(Inches(panel_x), Inches(iy + ih + 0.02), Inches(panel_w), Inches(0.3))
                    tf_c = cb.text_frame
                    tf_c.word_wrap = True
                    tf_c.margin_left = 0
                    tf_c.margin_right = 0
                    tf_c.margin_top = 0
                    tf_c.margin_bottom = 0
                    p_c = tf_c.paragraphs[0]
                    p_c.alignment = PP_ALIGN.CENTER
                    set_paragraph_text_clean(
                        p_c, cap, font_size_pt=11.0, is_italic=True,
                        color=theme.get("subtitle_color", RGBColor(0, 0, 0)),
                        font_name=theme["font_body"]
                    )

    return slide


def add_metrics_summary_slide(
    prs: PresentationType,
    title: str,
    cards: List[Dict[str, Any]],
    summary_bullets: Optional[List[Any]] = None,
    subtitle: Optional[str] = None,
    banner: Optional[str] = None,
    theme_name: str = "corporate_blue",
    card_font_size: float = 13.5,
    summary_font_size: float = 13.5,
    target_slide: Optional[Any] = None,
) -> Any:
    """
    Creates a Multi-Column Technical Metrics / Parameters Slide:
    - 2, 3, or 4 side-by-side parameter cards with bold metric rows.
    - Optional full-width bottom summary banner for system-level guarantees.
    """
    theme = get_theme(theme_name)
    slide = _prepare_slide(prs, theme, target_slide=target_slide)
    top_y = _add_slide_header(slide, title, subtitle, theme, banner=banner)

    total_w = 11.733
    start_x = 0.8
    content_top = top_y + 0.08
    total_h = max(3.5, 7.0 - content_top)

    has_summary = bool(summary_bullets)
    cards_h = total_h * 0.70 if has_summary else total_h
    summary_h = total_h * 0.25 if has_summary else 0.0
    summary_top = content_top + cards_h + (total_h * 0.05)

    # Column Cards
    count = max(1, min(len(cards), 4))
    gap = 0.20
    card_w = (total_w - (count - 1) * gap) / count

    for i, cdata in enumerate(cards[:count]):
        cx = start_x + i * (card_w + gap)
        card = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE,
            Inches(cx), Inches(content_top), Inches(card_w), Inches(cards_h)
        )
        card.fill.solid()
        card.fill.fore_color.rgb = theme["card_bg"]
        card.line.color.rgb = theme["card_border"]
        card.line.width = Pt(1.2)

        tf = card.text_frame
        tf.word_wrap = True
        tf.vertical_anchor = MSO_ANCHOR.TOP
        tf.margin_left = Inches(0.18)
        tf.margin_right = Inches(0.18)
        tf.margin_top = Inches(0.18)
        tf.margin_bottom = Inches(0.12)
        tf.clear()

        # Card Title
        p0 = tf.paragraphs[0]
        p0.space_after = Pt(8)
        set_paragraph_text_clean(
            p0, cdata.get("title", f"Module {i+1}"),
            font_size_pt=(17 if count <= 3 else 15),
            is_bold=True,
            color=theme.get("accent_color", RGBColor(0, 81, 226)),
            font_name=theme["font_heading"]
        )

        # Metrics / rows
        metrics = cdata.get("metrics", cdata.get("bullets", []))
        for m in metrics:
            p = tf.add_paragraph()
            _format_bullet_paragraph(p, m, theme, font_size_pt=card_font_size)

    # Bottom Summary Banner (if specified)
    if has_summary:
        b_card = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE,
            Inches(start_x), Inches(summary_top), Inches(total_w), Inches(summary_h)
        )
        b_card.fill.solid()
        b_card.fill.fore_color.rgb = theme["accent_light"]
        b_card.line.color.rgb = theme["accent_color"]
        b_card.line.width = Pt(1.2)

        b_tf = b_card.text_frame
        b_tf.word_wrap = True
        b_tf.vertical_anchor = MSO_ANCHOR.TOP
        b_tf.margin_left = Inches(0.24)
        b_tf.margin_right = Inches(0.24)
        b_tf.margin_top = Inches(0.12)
        b_tf.margin_bottom = Inches(0.10)
        b_tf.clear()

        first = True
        for b_item in summary_bullets:
            p = b_tf.paragraphs[0] if first else b_tf.add_paragraph()
            first = False
            _format_bullet_paragraph(p, b_item, theme, font_size_pt=summary_font_size)

    return slide


def add_thank_you_slide(
    prs: PresentationType,
    title: str = "THANK YOU FOR LISTENING!",
    subtitle: Optional[str] = None,
    presenter: Optional[str] = None,
    advisor: Optional[str] = None,
    qa_text: Optional[str] = "Questions & Answers Session",
    theme_name: str = "corporate_blue",
    target_slide: Optional[Any] = None,
) -> Any:
    """Creates a polished closing slide (Thank You & Q&A)."""
    theme = get_theme(theme_name)
    slide = _prepare_slide(prs, theme, target_slide=target_slide)

    card = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(1.64), Inches(1.64), Inches(10.05), Inches(4.60)
    )
    card.fill.solid()
    card.fill.fore_color.rgb = theme["card_bg"]
    card.line.color.rgb = theme["card_border"]
    card.line.width = Pt(1.5)

    tf = card.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.TOP
    tf.margin_top = Inches(0.4)
    tf.clear()

    p1 = tf.paragraphs[0]
    p1.alignment = PP_ALIGN.CENTER
    p1.space_after = Pt(16)
    set_paragraph_text_clean(
        p1, title, font_size_pt=36, is_bold=True,
        color=theme.get("accent_color", RGBColor(0, 81, 226)),
        font_name=theme["font_heading"]
    )

    if subtitle:
        p2 = tf.add_paragraph()
        p2.alignment = PP_ALIGN.CENTER
        p2.space_after = Pt(20)
        set_paragraph_text_clean(
            p2, subtitle, font_size_pt=22, is_bold=True,
            color=theme.get("title_color", RGBColor(0, 0, 0)),
            font_name=theme["font_heading"]
        )

    if presenter:
        p3 = tf.add_paragraph()
        p3.alignment = PP_ALIGN.CENTER
        p3.space_after = Pt(8)
        set_paragraph_text_clean(
            p3, presenter, font_size_pt=21, is_bold=False,
            color=theme.get("text_color", RGBColor(0, 0, 0)),
            font_name=theme["font_body"]
        )

    if advisor:
        p4 = tf.add_paragraph()
        p4.alignment = PP_ALIGN.CENTER
        p4.space_after = Pt(20)
        set_paragraph_text_clean(
            p4, advisor, font_size_pt=21, is_bold=False,
            color=theme.get("text_color", RGBColor(0, 0, 0)),
            font_name=theme["font_body"]
        )

    if qa_text:
        p5 = tf.add_paragraph()
        p5.alignment = PP_ALIGN.CENTER
        set_paragraph_text_clean(
            p5, qa_text, font_size_pt=25, is_bold=True, is_italic=True,
            color=theme.get("accent_color", RGBColor(0, 81, 226)),
            font_name=theme["font_heading"]
        )

    return slide


def add_kpi_cards_slide(
    prs: PresentationType,
    title: str,
    cards: List[Dict[str, str]],
    subtitle: Optional[str] = None,
    banner: Optional[str] = None,
    theme_name: str = "corporate_blue",
    target_slide: Optional[Any] = None,
) -> Any:
    """Creates a KPI / Feature Grid slide with 2, 3, or 4 visual metric cards."""
    theme = get_theme(theme_name)
    slide = _prepare_slide(prs, theme, target_slide=target_slide)
    top_y = _add_slide_header(slide, title, subtitle, theme, banner=banner)

    count = max(1, min(len(cards), 4))
    total_w = 11.733
    gap = 0.3
    card_w = (total_w - (count - 1) * gap) / count
    card_top = top_y + 0.2
    card_h = min(4.4, 7.0 - card_top)

    for i, cdata in enumerate(cards[:4]):
        card_x = 0.8 + i * (card_w + gap)

        card = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE,
            Inches(card_x), Inches(card_top), Inches(card_w), Inches(card_h)
        )
        card.fill.solid()
        card.fill.fore_color.rgb = theme["card_bg"]
        card.line.color.rgb = theme["card_border"]
        card.line.width = Pt(1.5)

        tf = card.text_frame
        tf.word_wrap = True
        tf.margin_left = Inches(0.3)
        tf.margin_right = Inches(0.3)
        tf.margin_top = Inches(0.4)

        p_title = tf.paragraphs[0]
        p_title.space_after = Pt(12)
        set_paragraph_text_clean(
            p_title, cdata.get("title", f"Metric {i+1}").upper(),
            font_size_pt=12, is_bold=True,
            color=theme.get("subtitle_color", RGBColor(0, 0, 0)),
            font_name=theme["font_heading"]
        )

        val_str = cdata.get("value", "0")
        p_val = tf.add_paragraph()
        p_val.space_after = Pt(10)
        set_paragraph_text_clean(
            p_val, val_str, font_size_pt=36, is_bold=True,
            color=theme.get("accent_color", RGBColor(0, 81, 226)),
            font_name=theme["font_heading"]
        )

        desc_str = cdata.get("description", cdata.get("desc", ""))
        if desc_str:
            p_desc = tf.add_paragraph()
            set_paragraph_text_clean(
                p_desc, desc_str, font_size_pt=13, is_bold=False,
                color=theme.get("text_color", RGBColor(0, 0, 0)),
                font_name=theme["font_body"]
            )

    return slide


def add_grid_cards_slide(
    prs: PresentationType,
    title: str,
    cards: List[Dict[str, str]],
    subtitle: Optional[str] = None,
    banner: Optional[str] = None,
    theme_name: str = "thesis_blue",
    target_slide: Optional[Any] = None,
) -> Any:
    """Creates a clean Grid Cards slide (e.g. 4x2 grid or 2x2 grid) for Agendas and Chapter overviews."""
    theme = get_theme(theme_name)
    slide = _prepare_slide(prs, theme, target_slide=target_slide)
    content_top = _add_slide_header(slide, title, subtitle, theme, banner=banner)

    n = len(cards)
    cols = 2 if n <= 4 else (3 if n <= 6 else 4)
    rows = (n + cols - 1) // cols

    total_w = 11.733
    total_h = min(5.0, 7.0 - content_top - 0.2)
    gap_x = 0.25
    gap_y = 0.22
    card_w = (total_w - (cols - 1) * gap_x) / cols
    card_h = (total_h - (rows - 1) * gap_y) / rows

    for idx, cdata in enumerate(cards):
        r = idx // cols
        c = idx % cols
        cx = 0.8 + c * (card_w + gap_x)
        cy = content_top + 0.12 + r * (card_h + gap_y)

        card = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE,
            Inches(cx), Inches(cy), Inches(card_w), Inches(card_h)
        )
        card.fill.solid()
        card.fill.fore_color.rgb = theme["card_bg"]
        card.line.color.rgb = theme["card_border"]
        card.line.width = Pt(1.2)

        tf = card.text_frame
        tf.word_wrap = True
        tf.margin_left = Inches(0.18)
        tf.margin_right = Inches(0.18)
        tf.margin_top = Inches(0.14)
        tf.margin_bottom = Inches(0.14)

        p_t = tf.paragraphs[0]
        set_paragraph_text_clean(
            p_t, cdata.get("title", f"Item {idx+1}"),
            font_size_pt=(13 if cols >= 4 else 15), is_bold=True,
            color=theme.get("title_color", RGBColor(0, 0, 0)),
            font_name=theme["font_heading"]
        )

        desc = cdata.get("description", cdata.get("desc", ""))
        if desc:
            p_d = tf.add_paragraph()
            p_d.space_before = Pt(3)
            set_paragraph_text_clean(
                p_d, desc,
                font_size_pt=(10 if cols >= 4 else 12), is_bold=False,
                color=theme.get("text_color", RGBColor(0, 0, 0)),
                font_name=theme["font_body"]
            )

    return slide


def add_table_slide(
    prs: PresentationType,
    title: str,
    headers: List[str],
    rows: List[List[Any]],
    subtitle: Optional[str] = None,
    banner: Optional[str] = None,
    theme_name: str = "corporate_blue",
    target_slide: Optional[Any] = None,
) -> Any:
    """Creates a native OpenXML Table slide with styled headers (PPTX_INV_06)."""
    theme = get_theme(theme_name)
    slide = _prepare_slide(prs, theme, target_slide=target_slide)
    top_y = _add_slide_header(slide, title, subtitle, theme, banner=banner)

    num_rows = len(rows) + 1
    num_cols = len(headers)
    table_shape = slide.shapes.add_table(
        num_rows, num_cols,
        Inches(0.8), Inches(top_y + 0.1), Inches(11.733), Inches(min(5.0, 0.45 * num_rows))
    )
    tbl = table_shape.table

    # Format Header Row
    for col_idx, h_text in enumerate(headers):
        cell = tbl.cell(0, col_idx)
        cell.fill.solid()
        cell.fill.fore_color.rgb = theme["title_color"]
        p = cell.text_frame.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        set_paragraph_text_clean(
            p, h_text, font_size_pt=14, is_bold=True,
            color=RGBColor(255, 255, 255),
            font_name=theme["font_heading"]
        )

    # Format Data Rows
    for row_idx, r_data in enumerate(rows):
        is_even = (row_idx % 2 == 0)
        row_bg = theme["card_bg"] if is_even else theme["accent_light"]

        for col_idx in range(num_cols):
            cell = tbl.cell(row_idx + 1, col_idx)
            cell.fill.solid()
            cell.fill.fore_color.rgb = row_bg
            val = r_data[col_idx] if col_idx < len(r_data) else ""
            p = cell.text_frame.paragraphs[0]
            if col_idx > 0 and str(val).replace('.', '', 1).isdigit():
                p.alignment = PP_ALIGN.RIGHT
            set_paragraph_text_clean(
                p, str(val), font_size_pt=12, is_bold=False,
                color=theme.get("text_color", RGBColor(0, 0, 0)),
                font_name=theme["font_body"]
            )
    return slide


def paginate_slide_specs(
    slides_data: List[Dict[str, Any]],
    max_bullets_threshold: int = 7,
    max_chars_threshold: int = 550,
) -> List[Dict[str, Any]]:
    """
    Auto-paginates content slides that exceed safe presentation density.
    If a content slide has > 7 bullets or high text volume (> 550 chars / > 9 estimated lines),
    it is automatically split into consecutive slides with the EXACT SAME title (no suffix).
    """
    paginated: List[Dict[str, Any]] = []
    for s in slides_data:
        stype = s.get("type", "content")
        # Only paginate generative content slides (not in-place targeted slides unless explicitly requested)
        if stype == "content" and s.get("auto_paginate", True) and s.get("slide_index") is None:
            bullets = s.get("bullets", [])
            if isinstance(bullets, list) and len(bullets) > 0:
                total_chars = sum(
                    len(str(b.get("text", b) if isinstance(b, dict) else (b[1] if isinstance(b, (tuple, list)) and len(b) > 1 else b)))
                    for b in bullets
                )

                # Check overflow condition: > 7 bullets OR (>= 5 bullets and high char count)
                if len(bullets) > max_bullets_threshold or (len(bullets) >= 5 and total_chars > max_chars_threshold):
                    # Balance bullets across chunks (e.g. 10 -> [5, 5]; 8 -> [4, 4]; 12 -> [6, 6])
                    num_splits = max(2, (len(bullets) + max_bullets_threshold - 1) // max_bullets_threshold)
                    chunk_size = (len(bullets) + num_splits - 1) // num_splits
                    chunk_size = max(3, min(chunk_size, 6))

                    chunks = [bullets[i:i + chunk_size] for i in range(0, len(bullets), chunk_size)]
                    for chunk in chunks:
                        new_s = dict(s)
                        new_s["bullets"] = chunk
                        # Exact same title (no suffix per user requirement)
                        paginated.append(new_s)
                    continue

        paginated.append(s)
    return paginated


# ---------------------------------------------------------------------------
# High-Level Spec Dispatcher
# ---------------------------------------------------------------------------

def write_pptx_from_spec(
    spec: Dict[str, Any],
    output_path: str,
    theme_name: str = "corporate_blue",
    template_path: Optional[str] = None,
) -> str:
    """
    Compiles a structured JSON Slide Spec dictionary into a finished .pptx file.
    Supports in_place_template mode for mutating existing slides in-place.
    """
    applied_theme = spec.get("theme", theme_name)
    theme = get_theme(applied_theme)
    tpl = template_path or spec.get("template")
    prs = init_presentation(template_path=tpl)

    is_in_place = spec.get("mode") == "in_place_template" or bool(tpl and spec.get("in_place", False))
    raw_slides_data = spec.get("slides", [])
    slides_data = paginate_slide_specs(raw_slides_data) if not is_in_place else raw_slides_data

    for s_idx, s in enumerate(slides_data):
        stype = s.get("type", "content")
        title = s.get("title", "Untitled Slide")
        subtitle = s.get("subtitle")
        banner = s.get("banner")
        slide_num = s.get("slide_number", s.get("number", s_idx + 1))

        # Check for skip/keep
        if stype in ("keep", "skip", "pass"):
            continue

        target_slide = None
        if is_in_place:
            target_idx = s.get("slide_index")
            if target_idx is not None:
                # 1-indexed slide_index
                actual_idx = target_idx - 1
            else:
                actual_idx = s_idx
            if 0 <= actual_idx < len(prs.slides):
                target_slide = prs.slides[actual_idx]

        # Dispatch to slide builders
        slide_obj = None
        if stype == "title":
            slide_obj = add_title_slide(
                prs,
                title=title,
                subtitle=subtitle,
                presenter=s.get("presenter"),
                date_str=s.get("date"),
                theme_name=applied_theme,
                target_slide=target_slide,
            )
        elif stype in ("chapter", "chapter_divider"):
            slide_obj = add_chapter_slide(
                prs,
                chapter_num=s.get("chapter_num", s.get("number", "CHAPTER")),
                title=title,
                subtitle=subtitle,
                theme_name=applied_theme,
                target_slide=target_slide,
            )
        elif stype == "content":
            slide_obj = add_content_slide(
                prs,
                title=title,
                bullets=s.get("bullets", []),
                subtitle=subtitle,
                banner=banner,
                theme_name=applied_theme,
                bullet_font_size=float(s.get("bullet_font_size", 16.5)),
                target_slide=target_slide,
            )
        elif stype == "two_column":
            slide_obj = add_two_column_slide(
                prs,
                title=title,
                left_bullets=s.get("left_bullets", []),
                right_bullets=s.get("right_bullets", []),
                subtitle=subtitle,
                banner=banner,
                theme_name=applied_theme,
                target_slide=target_slide,
            )
        elif stype == "diagram":
            slide_obj = add_diagram_slide(
                prs,
                title=title,
                image_path=s.get("image_path", ""),
                caption=s.get("caption"),
                subtitle=subtitle,
                banner=banner,
                theme_name=applied_theme,
                target_slide=target_slide,
            )
        elif stype in ("split_diagram", "split_card_diagram"):
            slide_obj = add_split_diagram_slide(
                prs,
                title=title,
                bullets=s.get("bullets", []),
                image_path=s.get("image_path", ""),
                caption=s.get("caption"),
                subtitle=subtitle,
                banner=banner,
                theme_name=applied_theme,
                card_width_ratio=float(s.get("card_width_ratio", 0.50)),
                bullet_font_size=float(s.get("bullet_font_size", 16.0)),
                target_slide=target_slide,
            )
        elif stype in ("dual_diagram", "split_card_dual_diagram"):
            images = s.get("images", [])
            img1 = s.get("img1_path", s.get("image_path_1", images[0].get("path") if len(images) > 0 else ""))
            cap1 = s.get("cap1", s.get("caption_1", images[0].get("caption") if len(images) > 0 else ""))
            img2 = s.get("img2_path", s.get("image_path_2", images[1].get("path") if len(images) > 1 else ""))
            cap2 = s.get("cap2", s.get("caption_2", images[1].get("caption") if len(images) > 1 else ""))
            slide_obj = add_dual_diagram_slide(
                prs,
                title=title,
                bullets=s.get("bullets", []),
                img1_path=img1,
                cap1=cap1,
                img2_path=img2,
                cap2=cap2,
                subtitle=subtitle,
                banner=banner,
                orientation=s.get("orientation", "vertical"),
                theme_name=applied_theme,
                card_width_ratio=float(s.get("card_width_ratio", 0.50)),
                bullet_font_size=float(s.get("bullet_font_size", 16.0)),
                target_slide=target_slide,
            )
        elif stype in ("metrics_summary", "metrics_cards", "multi_card_summary"):
            slide_obj = add_metrics_summary_slide(
                prs,
                title=title,
                cards=s.get("cards", []),
                summary_bullets=s.get("summary_bullets", s.get("summary", None)),
                subtitle=subtitle,
                banner=banner,
                theme_name=applied_theme,
                card_font_size=float(s.get("card_font_size", 13.5)),
                summary_font_size=float(s.get("summary_font_size", 13.5)),
                target_slide=target_slide,
            )
        elif stype in ("thank_you", "closing", "qa"):
            slide_obj = add_thank_you_slide(
                prs,
                title=title or "THANK YOU FOR LISTENING!",
                subtitle=subtitle,
                presenter=s.get("presenter"),
                advisor=s.get("advisor"),
                qa_text=s.get("qa_text", "Questions & Answers Session"),
                theme_name=applied_theme,
                target_slide=target_slide,
            )
        elif stype == "kpi_cards":
            slide_obj = add_kpi_cards_slide(
                prs,
                title=title,
                cards=s.get("cards", []),
                subtitle=subtitle,
                banner=banner,
                theme_name=applied_theme,
                target_slide=target_slide,
            )
        elif stype in ("grid_cards", "cards"):
            slide_obj = add_grid_cards_slide(
                prs,
                title=title,
                cards=s.get("cards", []),
                subtitle=subtitle,
                banner=banner,
                theme_name=applied_theme,
                target_slide=target_slide,
            )
        elif stype == "table":
            slide_obj = add_table_slide(
                prs,
                title=title,
                headers=s.get("headers", []),
                rows=s.get("rows", []),
                subtitle=subtitle,
                banner=banner,
                theme_name=applied_theme,
                target_slide=target_slide,
            )

        if slide_obj is not None and is_in_place:
            _update_slide_number(slide_obj, slide_num, theme)

    # Trim extra slides if slide_count specified
    target_count = spec.get("slide_count")
    if target_count is not None and len(prs.slides) > target_count:
        while len(prs.slides) > target_count:
            idx_to_del = len(prs.slides) - 1
            slide_id = prs.slides._sldIdLst[idx_to_del]
            prs.slides._sldIdLst.remove(slide_id)

    return safe_save_pptx(prs, output_path)

