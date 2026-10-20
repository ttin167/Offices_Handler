"""
Bidirectional Markdown Converter for PDF, DOCX, and Markdown (.md)
Implements decoupled Document Architecture:
- Data Layer: Pure standard GFM Markdown (.md) containing 100% data fidelity (Headings, Bullets, Tables, Images).
- Presentation Layer: Clean YAML stylesheet (.style.yaml) acting like CSS for layout & styling.
- Round-trip compiler: (Data .md + Style .style.yaml) -> DOCX -> PDF.
"""

from __future__ import annotations
import os
import re
import sys
import yaml
from typing import Any, Callable, Dict, List, Optional, Tuple, Union

import pymupdf
import docx
import docx.text
import docx.text.paragraph
import docx.table
import docx.oxml
from docx.document import Document as DocumentType
from docx.shared import Pt, Inches, Length, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

NS_W = nsdecls("w")
TOC_LINE_REGEX = re.compile(r"^(.*?)(?:\.{2,}|…+|(?:\.\s+){2,}\.|\t)\s*(\d+)$")

try:
    from .smart_post_processor import calculate_heuristic_column_widths
except (ImportError, ValueError):
    from smart_post_processor import calculate_heuristic_column_widths

try:
    from .docx_math import add_math_to_paragraph
except (ImportError, ValueError):
    from docx_math import add_math_to_paragraph

try:
    from .docx_lists import ListContextTracker, apply_list_formatting
except (ImportError, ValueError):
    from docx_lists import ListContextTracker, apply_list_formatting

def _hex_to_rgb(hex_str: str) -> RGBColor:
    cleaned = hex_str.lstrip("#")
    if len(cleaned) == 6:
        try:
            r = int(cleaned[0:2], 16)
            g = int(cleaned[2:4], 16)
            b = int(cleaned[4:6], 16)
            return RGBColor(r, g, b)
        except ValueError:
            pass
    return RGBColor(30, 41, 59)


def _parse_length(val: Any, default_inches: float = 1.0) -> Union[Inches, Pt, Length]:
    if val is None:
        return Inches(default_inches)
    if isinstance(val, (int, float)):
        return Inches(float(val))
    val_s = str(val).strip().lower()
    if val_s.endswith("in") or val_s.endswith("inch") or val_s.endswith("inches"):
        num = re.sub(r"[a-z]", "", val_s).strip()
        try:
            return Inches(float(num))
        except ValueError:
            return Inches(default_inches)
    elif val_s.endswith("cm"):
        num = val_s.replace("cm", "").strip()
        try:
            return Inches(float(num) / 2.54)
        except ValueError:
            return Inches(default_inches)
    elif val_s.endswith("pt"):
        num = val_s.replace("pt", "").strip()
        try:
            return Pt(float(num))
        except ValueError:
            return Inches(default_inches)
    try:
        return Inches(float(val_s))
    except ValueError:
        return Inches(default_inches)


def extract_pdf_styling_metadata(doc: pymupdf.Document) -> Dict[str, Any]:
    """
    Scan PDF sample pages to detect ground truth geometry and styling:
    - Page dimensions & margins (top, bottom, left, right).
    - Dominant body font family & size.
    - Prominent heading color (e.g. Red #C00000).
    - Dominant table header background fill color (e.g. Peach #FFE8E0).
    """
    if len(doc) == 0:
        return {}

    first_page = doc[0]
    p_w, p_h = first_page.rect.width, first_page.rect.height
    is_letter = abs(p_w - 612) < 30 and abs(p_h - 792) < 30
    page_size = "Letter" if is_letter else "A4"

    sample_limit = min(20, len(doc))
    lefts, rights, tops, bottoms = [], [], [], []
    fonts: Dict[str, int] = {}
    colors: Dict[str, int] = {}
    heading_colors: Dict[str, int] = {}
    fill_counts: Dict[str, int] = {}

    for p_idx in range(sample_limit):
        p = doc[p_idx]

        # Scan text blocks for margins
        blocks = p.get_text("blocks")
        for b in blocks:
            if b[1] > 45 and b[3] < p_h - 45 and (b[2] - b[0]) > 40:
                lefts.append(b[0])
                rights.append(p_w - b[2])
                tops.append(b[1])
                bottoms.append(p_h - b[3])

        # Scan text spans for fonts & colors
        td = p.get_text("dict")
        for b in td.get("blocks", []):
            if "lines" not in b:
                continue
            for l in b["lines"]:
                for s in l.get("spans", []):
                    txt = s.get("text", "").strip()
                    if not txt:
                        continue
                    font_base = s.get("font", "Calibri").split(",")[0].split("-")[0]
                    sz = round(s.get("size", 11.0), 1)
                    c = s.get("color", 0)
                    hex_c = f"#{c:06X}" if isinstance(c, int) else "#000000"

                    fonts[font_base] = fonts.get(font_base, 0) + len(txt)
                    colors[hex_c] = colors.get(hex_c, 0) + len(txt)

                    if sz >= 14 or "Bold" in s.get("font", ""):
                        if hex_c not in ("#000000", "#FFFFFF", "#202124", "#333333"):
                            heading_colors[hex_c] = heading_colors.get(hex_c, 0) + len(txt)

        # Scan drawings for table header vector fills
        for d in p.get_drawings():
            fill = d.get("fill")
            if fill:
                r, g, b_val = [int(c_val * 255) for c_val in fill]
                hex_fill = f"#{r:02X}{g:02X}{b_val:02X}"
                if hex_fill not in ("#000000", "#FFFFFF"):
                    fill_counts[hex_fill] = fill_counts.get(hex_fill, 0) + 1

    def to_in_str(val_list: List[float], default_in: float) -> str:
        if not val_list:
            return f"{default_in:.2f}in"
        val_in = round(min(val_list) / 72.0, 2)
        val_in = max(0.40, min(1.25, val_in))
        return f"{val_in:.2f}in"

    margin_left = to_in_str(lefts, 1.0)
    margin_right = to_in_str(rights, 0.55)
    margin_top = to_in_str(tops, 0.63)
    margin_bottom = to_in_str(bottoms, 0.87)

    dom_font = sorted(fonts.items(), key=lambda x: x[1], reverse=True)[0][0] if fonts else "Calibri"
    dom_color = sorted(colors.items(), key=lambda x: x[1], reverse=True)[0][0] if colors else "#000000"
    heading_color = sorted(heading_colors.items(), key=lambda x: x[1], reverse=True)[0][0] if heading_colors else "#C00000"
    table_header_bg = sorted(fill_counts.items(), key=lambda x: x[1], reverse=True)[0][0] if fill_counts else "#FFE8E0"

    return {
        "geometry": {
            "page_size": page_size,
            "margin": {
                "top": margin_top,
                "bottom": margin_bottom,
                "left": margin_left,
                "right": margin_right
            }
        },
        "typography": {
            "font_family": f"{dom_font}, Arial, sans-serif",
            "font_size": "11pt",
            "body_color": dom_color,
            "line_spacing": 1.15
        },
        "headings": {
            "color": heading_color,
            "h1_size": "16pt",
            "h2_size": "13pt",
            "h3_size": "12pt",
            "h4_size": "11pt"
        },
        "tables": {
            "header_background": table_header_bg,
            "border_color": "#000000",
            "cell_padding": "6pt"
        },
        "lists": {
            "bullet_glyph": "•",
            "space_before": "2pt",
            "space_after": "2pt"
        }
    }


def extract_docx_styling_metadata(doc: DocumentType) -> Dict[str, Any]:
    """Extract geometry and styling schema from an existing DOCX document."""
    sec = doc.sections[0]
    p_w = sec.page_width.inches if sec.page_width else 8.5
    p_h = sec.page_height.inches if sec.page_height else 11.0
    page_size = "Letter" if abs(p_w - 8.5) < 0.5 and abs(p_h - 11.0) < 0.5 else "A4"

    top_m = f"{sec.top_margin.inches:.2f}in" if sec.top_margin else "0.63in"
    bottom_m = f"{sec.bottom_margin.inches:.2f}in" if sec.bottom_margin else "0.87in"
    left_m = f"{sec.left_margin.inches:.2f}in" if sec.left_margin else "1.00in"
    right_m = f"{sec.right_margin.inches:.2f}in" if sec.right_margin else "0.55in"

    # Font & Color detection
    fonts: Dict[str, int] = {}
    colors: Dict[str, int] = {}
    heading_colors: Dict[str, int] = {}

    for p in doc.paragraphs:
        for r in p.runs:
            txt = r.text.strip()
            if not txt:
                continue
            if r.font.name:
                fonts[r.font.name] = fonts.get(r.font.name, 0) + len(txt)
            if r.font.color and r.font.color.rgb:
                hex_c = f"#{r.font.color.rgb}"
                colors[hex_c] = colors.get(hex_c, 0) + len(txt)
                style_name = (p.style.name if p.style and p.style.name else "")
                if style_name.startswith("Heading") or r.bold:
                    if hex_c.upper() not in ("#000000", "#FFFFFF", "#1E293B", "#333333", "#202124"):
                        heading_colors[hex_c.upper()] = heading_colors.get(hex_c.upper(), 0) + len(txt)

    dom_font = sorted(fonts.items(), key=lambda x: x[1], reverse=True)[0][0] if fonts else "Calibri"
    dom_color = sorted(colors.items(), key=lambda x: x[1], reverse=True)[0][0] if colors else "#000000"
    heading_color = sorted(heading_colors.items(), key=lambda x: x[1], reverse=True)[0][0] if heading_colors else "#C00000"

    # Table header fill detection
    table_header_bg = "#FFE8E0"
    for t in doc.tables:
        if t.rows:
            tcPr = t.rows[0].cells[0]._tc.get_or_add_tcPr()
            shd = tcPr.find(qn("w:shd"))
            if shd is not None:
                fill_val = shd.get(qn("w:fill"))
                if fill_val and fill_val.lower() not in ("auto", "none", "ffffff", "000000"):
                    table_header_bg = f"#{fill_val.upper()}"
                    break

    return {
        "geometry": {
            "page_size": page_size,
            "margin": {
                "top": top_m,
                "bottom": bottom_m,
                "left": left_m,
                "right": right_m
            }
        },
        "typography": {
            "font_family": f"{dom_font}, Arial, sans-serif",
            "font_size": "11pt",
            "body_color": dom_color,
            "line_spacing": 1.15
        },
        "headings": {
            "color": heading_color,
            "h1_size": "16pt",
            "h2_size": "13pt",
            "h3_size": "12pt",
            "h4_size": "11pt"
        },
        "tables": {
            "header_background": table_header_bg,
            "border_color": "#000000",
            "cell_padding": "6pt"
        },
        "lists": {
            "bullet_glyph": "•",
            "space_before": "2pt",
            "space_after": "2pt"
        }
    }


# ==============================================================================
# 1. DOCX -> Decoupled (Data .md + Style .style.yaml)
# ==============================================================================
def docx_to_decoupled(
    docx_path: str,
    md_path: Optional[str] = None,
    style_yaml_path: Optional[str] = None,
    progress_callback: Optional[Callable[[float, str], None]] = None
) -> Tuple[bool, str, str]:
    """
    Extracts a pristine DOCX into two decoupled files:
    1. [name].md: Pure GFM Markdown containing 100% semantic content (no styling tags).
    2. [name].style.yaml: Structured layout & typography schema (CSS equivalent).
    """
    if not os.path.isfile(docx_path):
        return False, f"File không tồn tại: {docx_path}", ""

    stem = os.path.splitext(docx_path)[0]
    if not md_path:
        md_path = stem + ".md"
    if not style_yaml_path:
        style_yaml_path = stem + ".style.yaml"

    out_dir = os.path.dirname(os.path.abspath(md_path))
    base_name = os.path.splitext(os.path.basename(md_path))[0]
    assets_dir = os.path.join(out_dir, f"{base_name}_assets")
    os.makedirs(assets_dir, exist_ok=True)

    doc = docx.Document(docx_path)

    # 1. Extract and Save Style YAML
    style_meta = extract_docx_styling_metadata(doc)
    with open(style_yaml_path, "w", encoding="utf-8") as f_yaml:
        yaml.dump(style_meta, f_yaml, sort_keys=False, allow_unicode=True)

    # 2. Extract Images from relationships
    img_map = {}
    img_idx = 0
    for rel in doc.part.rels.values():
        if "image" in rel.target_ref:
            img_idx += 1
            img_ext = os.path.splitext(rel.target_ref)[1] or ".png"
            img_name = f"docx_img_{img_idx}{img_ext}"
            img_disk = os.path.join(assets_dir, img_name)
            with open(img_disk, "wb") as f_im:
                f_im.write(rel.target_part.blob)
            img_map[rel.target_ref] = f"{base_name}_assets/{img_name}"

    md_lines: List[str] = []
    total_elements = len(doc.paragraphs) + len(doc.tables)

    # Iterate elements preserving order
    # Extract document body elements in XML sequence
    doc_body = doc._element.body
    curr_idx = 0

    for child in doc_body:
        curr_idx += 1
        if progress_callback and total_elements > 0:
            percent = (curr_idx / total_elements) * 90.0
            progress_callback(percent, f"Đang bóc tách Markdown thuần ({curr_idx}/{total_elements})...")

        tag = child.tag.split("}")[-1]

        if tag == "p":
            p = docx.text.paragraph.Paragraph(child, doc)
            txt = p.text.strip()

            # Check if this paragraph contains embedded images/drawings
            blip_nodes = child.xpath(".//a:blip") or child.findall(".//{http://schemas.openxmlformats.org/drawingml/2006/main}blip")
            if not blip_nodes:
                blip_nodes = child.findall(".//{urn:schemas-microsoft-com:vml}imagedata")

            p_images = []
            for b in blip_nodes:
                r_id = (
                    b.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed")
                    or b.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
                    or b.get(qn("r:id"))
                    or b.get(qn("r:embed"))
                )
                if r_id and r_id in img_map:
                    p_images.append(img_map[r_id])

            if p_images:
                for img_rel in p_images:
                    md_lines.append(f"\n![Image]({img_rel})\n")
                if not txt:
                    continue

            if not txt:
                continue

            style_name = (p.style.name if p.style and p.style.name else "").lower()
            pPr = p._p.pPr
            has_num = pPr is not None and pPr.find(qn("w:numPr")) is not None
            has_tabs = pPr is not None and pPr.find(qn("w:tabs")) is not None

            # Check if TOC line with tab stop
            if has_tabs or "\t" in txt or TOC_LINE_REGEX.match(txt):
                m_t = TOC_LINE_REGEX.match(txt)
                if m_t:
                    md_lines.append(f"{m_t.group(1).strip()} ... {m_t.group(2).strip()}")
                    continue
                elif "\t" in txt:
                    parts = txt.split("\t", 1)
                    md_lines.append(f"{parts[0].strip()} ... {parts[1].strip()}")
                    continue

            # Headings
            if "heading 1" in style_name or re.match(r"^[IVXLCDM]+\.\s+[A-Z]", txt):
                clean_h = txt.strip()
                md_lines.append(f"\n# {clean_h}\n")
            elif "heading 2" in style_name or re.match(r"^\d+\.\s+[A-Z]", txt):
                clean_h = txt.strip()
                md_lines.append(f"\n## {clean_h}\n")
            elif "heading 3" in style_name or re.match(r"^\d+\.\d+\s+[A-Z]", txt):
                clean_h = txt.strip()
                md_lines.append(f"\n### {clean_h}\n")
            elif "heading 4" in style_name or re.match(r"^\d+\.\d+\.\d+\s+[A-Z]", txt):
                clean_h = txt.strip()
                md_lines.append(f"\n#### {clean_h}\n")
            # Bullet list items
            elif has_num or "bullet" in style_name or txt.startswith("•") or txt.startswith("+") or txt.startswith("-"):
                item_clean = re.sub(r"^[\uf06c\uf0b7•+\-*]\s*", "", txt).strip()
                md_lines.append(f"- {item_clean}")
            # Captions
            elif txt.startswith("Table ") or txt.startswith("Figure "):
                md_lines.append(f"*{txt}*\n")
            else:
                md_lines.append(f"{txt}\n")

        elif tag == "tbl":
            t = docx.table.Table(child, doc)
            if not t.rows:
                continue

            cleaned_rows: List[List[str]] = []
            for row in t.rows:
                row_cells = [c.text.strip().replace("\n", " ") for c in row.cells]
                if any(c for c in row_cells):
                    cleaned_rows.append(row_cells)

            if not cleaned_rows:
                continue

            # Filter duplicate columns from merged cells
            num_cols = len(cleaned_rows[0])
            table_md = []
            # Header
            table_md.append("| " + " | ".join(cleaned_rows[0]) + " |")
            table_md.append("| " + " | ".join(["---"] * num_cols) + " |")
            for r in cleaned_rows[1:]:
                # Pad or trim to num_cols
                padded = r[:num_cols] + [""] * max(0, num_cols - len(r))
                table_md.append("| " + " | ".join(padded) + " |")

            md_lines.append("\n" + "\n".join(table_md) + "\n")

    with open(md_path, "w", encoding="utf-8") as f_md:
        f_md.write("\n".join(md_lines))

    if progress_callback:
        progress_callback(100.0, f"Hoàn tất bóc tách: {os.path.basename(md_path)} & {os.path.basename(style_yaml_path)}")

    return True, md_path, style_yaml_path


# ==============================================================================
# Helper: Diagram Rendering & Embedding for Markdown Code Blocks
# ==============================================================================
def _render_and_embed_mermaid_diagram(
    doc: DocumentType,
    code: str,
    assets_dir: str,
    font_family: str
) -> bool:
    """Render Mermaid code block into high-res PNG and embed into Word document."""
    try:
        try:
            from core.diagrams.legacy.mermaid_renderer import render_mermaid_to_png, calculate_aspect_dimensions
        except (ImportError, ValueError):
            from legacy_engines.mermaid_renderer import render_mermaid_to_png, calculate_aspect_dimensions

        import hashlib
        code_hash = hashlib.md5(code.encode("utf-8")).hexdigest()[:10]
        out_png = os.path.join(assets_dir, f"diag_mermaid_{code_hash}.png")

        if not (os.path.isfile(out_png) and os.path.getsize(out_png) > 0):
            res = render_mermaid_to_png({
                "mermaid_code": code,
                "output_path": out_png,
                "scale": 2,
            })
            if not res or not res.get("success"):
                return False

        if os.path.isfile(out_png):
            from PIL import Image
            with Image.open(out_png) as im:
                w_px, h_px = im.size
            w_cm, h_cm = calculate_aspect_dimensions(w_px, h_px, max_w_cm=14.0, max_h_cm=18.0)

            p_img = doc.add_paragraph()
            p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_img.paragraph_format.space_before = Pt(8)
            p_img.paragraph_format.space_after = Pt(2)
            r_img = p_img.add_run()
            r_img.add_picture(out_png, width=Inches(w_cm / 2.54))

            # Extract title if present in code
            t_match = re.search(r'title\s+["\']?([^"\']+)["\']?', code, re.IGNORECASE)
            cap_text = f"Sơ đồ Mermaid: {t_match.group(1).strip()}" if t_match else "Sơ đồ Mermaid"

            p_cap = doc.add_paragraph()
            p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r_cap = p_cap.add_run(cap_text)
            r_cap.font.name = font_family
            r_cap.font.size = Pt(9.5)
            r_cap.italic = True
            p_cap.paragraph_format.space_before = Pt(2)
            p_cap.paragraph_format.space_after = Pt(8)
            return True
    except Exception as exc:
        print(f"[WARN] Failed to render Mermaid diagram: {exc}", file=sys.stderr)
        return False
    return False


def _render_and_embed_plantuml_diagram(
    doc: DocumentType,
    code: str,
    assets_dir: str,
    font_family: str
) -> bool:
    """Render PlantUML code block into high-res PNG and embed into Word document."""
    try:
        try:
            from legacy_engines.plantuml_renderer import render_plantuml_to_png, calculate_aspect_dimensions
        except (ImportError, ValueError):
            from plantuml_renderer import render_plantuml_to_png, calculate_aspect_dimensions

        import hashlib
        code_hash = hashlib.md5(code.encode("utf-8")).hexdigest()[:10]
        out_png = os.path.join(assets_dir, f"diag_plantuml_{code_hash}.png")

        puml_code = code.strip()
        if not puml_code.startswith("@startuml"):
            puml_code = f"@startuml\n{puml_code}\n@enduml"

        if not (os.path.isfile(out_png) and os.path.getsize(out_png) > 0):
            res = render_plantuml_to_png({
                "plantuml_code": puml_code,
                "output_path": out_png,
                "scale": 2,
            })
            if not res or not res.get("success"):
                return False

        if os.path.isfile(out_png):
            from PIL import Image
            with Image.open(out_png) as im:
                w_px, h_px = im.size
            w_cm, h_cm = calculate_aspect_dimensions(w_px, h_px, max_w_cm=14.0, max_h_cm=18.0)

            p_img = doc.add_paragraph()
            p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_img.paragraph_format.space_before = Pt(8)
            p_img.paragraph_format.space_after = Pt(2)
            r_img = p_img.add_run()
            r_img.add_picture(out_png, width=Inches(w_cm / 2.54))

            t_match = re.search(r'title\s+([^\n]+)', code, re.IGNORECASE)
            cap_text = f"Sơ đồ PlantUML: {t_match.group(1).strip()}" if t_match else "Sơ đồ PlantUML"

            p_cap = doc.add_paragraph()
            p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r_cap = p_cap.add_run(cap_text)
            r_cap.font.name = font_family
            r_cap.font.size = Pt(9.5)
            r_cap.italic = True
            p_cap.paragraph_format.space_before = Pt(2)
            p_cap.paragraph_format.space_after = Pt(8)
            return True
    except Exception as exc:
        print(f"[WARN] Failed to render PlantUML diagram: {exc}", file=sys.stderr)
        return False
    return False


def _render_and_embed_mxgraph_diagram(
    doc: DocumentType,
    code: str,
    assets_dir: str,
    font_family: str
) -> bool:
    """Render Draw.io / mxGraph code block into high-res PNG and embed into Word document."""
    try:
        try:
            from core.diagrams.mxgraph import render_mxgraph_to_png
        except (ImportError, ValueError):
            from mxgraph_engine import render_mxgraph_to_png

        import hashlib
        code_hash = hashlib.md5(code.encode("utf-8")).hexdigest()[:10]
        out_png = os.path.join(assets_dir, f"diag_drawio_{code_hash}.png")

        if not (os.path.isfile(out_png) and os.path.getsize(out_png) > 0):
            res_png = render_mxgraph_to_png(code, out_png, scale=2)
            actual_png = res_png.get("png_path") if isinstance(res_png, dict) else (res_png if isinstance(res_png, str) else out_png)
            if not actual_png or not os.path.isfile(actual_png):
                return False
            out_png = actual_png

        if os.path.isfile(out_png):
            from PIL import Image
            with Image.open(out_png) as im:
                w_px, h_px = im.size
            max_w_cm = 14.0
            max_h_cm = 18.0
            if w_px > 0 and h_px > 0:
                scaled_h = (h_px / w_px) * max_w_cm
                if scaled_h <= max_h_cm:
                    w_cm, h_cm = max_w_cm, scaled_h
                else:
                    h_cm = max_h_cm
                    w_cm = min((w_px / h_px) * max_h_cm, max_w_cm)
            else:
                w_cm, h_cm = max_w_cm, max_h_cm

            p_img = doc.add_paragraph()
            p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_img.paragraph_format.space_before = Pt(8)
            p_img.paragraph_format.space_after = Pt(2)
            r_img = p_img.add_run()
            r_img.add_picture(out_png, width=Inches(w_cm / 2.54))

            p_cap = doc.add_paragraph()
            p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r_cap = p_cap.add_run("Sơ đồ Draw.io")
            r_cap.font.name = font_family
            r_cap.font.size = Pt(9.5)
            r_cap.italic = True
            p_cap.paragraph_format.space_before = Pt(2)
            p_cap.paragraph_format.space_after = Pt(8)
            return True
    except Exception as exc:
        print(f"[WARN] Failed to render Draw.io diagram: {exc}", file=sys.stderr)
        return False
    return False


def _flush_fenced_code_block(
    doc: DocumentType,
    lang: str,
    lines: List[str],
    assets_dir: str,
    font_family: str
) -> None:
    """Flush collected lines of a fenced code block, rendering diagrams if matched."""
    content = "\n".join(lines).strip()
    if not content:
        return

    clean_lang = (lang or "").strip().lower()
    rendered = False

    # Try unified diagram bridge first (Canvas, Mermaid, PlantUML, Draw.io)
    try:
        from core.diagrams.bridge import render_diagram_code_to_png, calculate_aspect_dimensions
    except (ImportError, ValueError):
        try:
            from diagram_bridge import render_diagram_code_to_png, calculate_aspect_dimensions
        except (ImportError, ValueError):
            render_diagram_code_to_png = None
            calculate_aspect_dimensions = None

    if render_diagram_code_to_png and clean_lang in (
        "mermaid", "mmd", "plantuml", "puml", "drawio", "mxgraph", "xml",
        "canvas", "spec", "diagram-spec", "diagram", "json"
    ):
        diag_res = render_diagram_code_to_png(content, clean_lang, base_dir=assets_dir)
        if diag_res and diag_res.get("png_path") and os.path.isfile(diag_res["png_path"]):
            out_png = diag_res["png_path"]
            w_px, h_px = diag_res["dimensions_px"]
            w_cm, h_cm = calculate_aspect_dimensions(w_px, h_px, max_w_cm=14.0, max_h_cm=18.0) if calculate_aspect_dimensions else (14.0, 10.0)

            p_img = doc.add_paragraph()
            p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_img.paragraph_format.space_before = Pt(8)
            p_img.paragraph_format.space_after = Pt(2)
            r_img = p_img.add_run()
            r_img.add_picture(out_png, width=Inches(w_cm / 2.54))

            engine_label = diag_res.get("engine", clean_lang).capitalize()
            title = diag_res.get("title")
            cap_text = f"Sơ đồ {engine_label}: {title}" if title else f"Sơ đồ {engine_label}"

            p_cap = doc.add_paragraph()
            p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r_cap = p_cap.add_run(cap_text)
            r_cap.font.name = font_family
            r_cap.font.size = Pt(9.5)
            r_cap.italic = True
            p_cap.paragraph_format.space_before = Pt(2)
            p_cap.paragraph_format.space_after = Pt(8)
            rendered = True

    # Fallback to legacy individual handlers if bridge didn't handle it
    if not rendered:
        if clean_lang in ("mermaid", "mmd"):
            rendered = _render_and_embed_mermaid_diagram(doc, content, assets_dir, font_family)
        elif clean_lang in ("plantuml", "puml"):
            rendered = _render_and_embed_plantuml_diagram(doc, content, assets_dir, font_family)
        elif clean_lang in ("drawio", "mxgraph"):
            rendered = _render_and_embed_mxgraph_diagram(doc, content, assets_dir, font_family)
        elif clean_lang in ("chart", "charts", "chart-spec", "chart_spec"):
            try:
                import json
                try:
                    from core.docx.charts import add_chart_to_doc
                except (ImportError, ValueError):
                    from docx_charts import add_chart_to_doc  # type: ignore

                spec_dict = None
                try:
                    spec_dict = json.loads(content)
                except Exception:
                    try:
                        spec_dict = yaml.safe_load(content)
                    except Exception:
                        pass

                if isinstance(spec_dict, dict):
                    add_chart_to_doc(doc, spec_dict)
                    rendered = True
            except Exception as chart_err:
                print(f"[WARN] Failed to render chart block: {chart_err}", file=sys.stderr)

    if not rendered:
        # Standard code block rendering
        for line in lines:
            p_code = doc.add_paragraph()
            r_code = p_code.add_run(line)
            r_code.font.name = "Consolas"
            r_code.font.size = Pt(9.5)
            p_code.paragraph_format.space_before = Pt(1)
            p_code.paragraph_format.space_after = Pt(1)
            p_code.paragraph_format.left_indent = Inches(0.2)


# ==============================================================================
# 2. Decoupled (Data .md + Style .style.yaml) -> DOCX
# ==============================================================================
def decoupled_to_docx(
    md_path: str,
    docx_path: Optional[str] = None,
    style_yaml_path: Optional[str] = None,
    progress_callback: Optional[Callable[[float, str], None]] = None
) -> Tuple[bool, str]:
    """
    Compiles pure Markdown data (.md) + YAML stylesheet (.style.yaml) into a high-fidelity DOCX.
    Automatically finds [stem].style.yaml in the same directory if not explicitly specified.
    """
    if not os.path.isfile(md_path):
        return False, f"File không tồn tại: {md_path}"

    stem = os.path.splitext(md_path)[0]
    if not docx_path:
        docx_path = stem + ".docx"

    # Auto-discover style yaml
    if not style_yaml_path:
        candidate = stem + ".style.yaml"
        if os.path.isfile(candidate):
            style_yaml_path = candidate
        else:
            # Check for candidate without .clean or similar
            candidate_base = re.sub(r"(_clean|_extracted|_perfect.*)$", "", stem, flags=re.IGNORECASE) + ".style.yaml"
            if os.path.isfile(candidate_base):
                style_yaml_path = candidate_base

    style_meta: Dict[str, Any] = {}
    if style_yaml_path and os.path.isfile(style_yaml_path):
        try:
            with open(style_yaml_path, "r", encoding="utf-8") as f_y:
                style_meta = yaml.safe_load(f_y) or {}
        except Exception:
            style_meta = {}

    with open(md_path, "r", encoding="utf-8") as f_md:
        body = f_md.read()

    # Fallback to embedded frontmatter if pure yaml not provided
    if not style_meta and body.startswith("---"):
        parts = body.split("---", 2)
        if len(parts) >= 3:
            try:
                style_meta = yaml.safe_load(parts[1]) or {}
            except Exception:
                pass
            body = parts[2]

    geom = style_meta.get("geometry", {})
    margins = geom.get("margin", {})
    typo = style_meta.get("typography", {})
    headings = style_meta.get("headings", {})
    tables_meta = style_meta.get("tables", {})

    font_family = typo.get("font_family", "Calibri").split(",")[0].strip().replace("'", "").replace('"', "")
    heading_hex = headings.get("color", "#C00000")
    body_hex = typo.get("body_color", "#000000")
    header_bg_hex = tables_meta.get("header_background", "#FFE8E0").lstrip("#")
    border_hex = tables_meta.get("border_color", "#000000").lstrip("#")

    doc = docx.Document()

    # 1. Page Margins
    sec = doc.sections[0]
    if isinstance(margins, dict):
        sec.top_margin = _parse_length(margins.get("top"), 0.63)
        sec.bottom_margin = _parse_length(margins.get("bottom"), 0.87)
        sec.left_margin = _parse_length(margins.get("left"), 1.00)
        sec.right_margin = _parse_length(margins.get("right"), 0.55)
    else:
        single_m = _parse_length(margins, 1.0)
        sec.top_margin = single_m
        sec.bottom_margin = single_m
        sec.left_margin = single_m
        sec.right_margin = single_m

    # Calculate Tab Stop position at right margin
    try:
        pw = sec.page_width or Inches(8.5)
        lm = sec.left_margin or Inches(1.0)
        rm = sec.right_margin or Inches(1.0)
        printable_w = int(pw) - int(lm) - int(rm)
        tab_pos_twips = int(printable_w / 635)
    except Exception:
        tab_pos_twips = 9360

    # 2. Configure Typography Styles
    normal_style = doc.styles["Normal"]
    normal_font = normal_style.font
    normal_font.name = font_family
    normal_font.size = Pt(11)
    normal_font.color.rgb = _hex_to_rgb(body_hex)
    normal_style.paragraph_format.line_spacing = 1.15
    normal_style.paragraph_format.space_before = Pt(0)
    normal_style.paragraph_format.space_after = Pt(4)

    h_color = _hex_to_rgb(heading_hex)
    for lvl, size in [
        ("Heading 1", Pt(16)),
        ("Heading 2", Pt(13)),
        ("Heading 3", Pt(12)),
        ("Heading 4", Pt(11)),
        ("Heading 5", Pt(11)),
        ("Heading 6", Pt(10.5)),
    ]:
        if lvl in doc.styles:
            h_st = doc.styles[lvl]
            h_st.font.name = font_family
            h_st.font.size = size
            h_st.font.bold = True
            h_st.font.color.rgb = h_color

    base_dir = os.path.dirname(os.path.abspath(md_path))
    docx_dir = os.path.dirname(os.path.abspath(docx_path))
    docx_stem = os.path.splitext(os.path.basename(docx_path))[0]
    assets_dir = os.path.join(docx_dir, f"{docx_stem}_assets")
    os.makedirs(assets_dir, exist_ok=True)

    lines = body.split("\n")
    total_lines = len(lines)

    in_table = False
    in_code_block = False
    code_block_lang = ""
    code_block_lines: List[str] = []
    in_signature_block = False
    table_buffer: List[List[str]] = []
    col_alignments: List[WD_ALIGN_PARAGRAPH] = []
    list_tracker = ListContextTracker(doc)

    def flush_table():
        nonlocal in_table, table_buffer, col_alignments
        if not table_buffer:
            return
        col_count = max((len(row) for row in table_buffer), default=0)
        if col_count == 0 or len(table_buffer) == 0:
            table_buffer.clear()
            col_alignments.clear()
            in_table = False
            return

        t = doc.add_table(rows=len(table_buffer), cols=col_count)
        t.autofit = False  # Enforce explicit heuristic column geometry

        # Calculate heuristic column widths based on Square Root Damped Weighting
        total_w_twips = tab_pos_twips or 9360
        col_widths_twips = calculate_heuristic_column_widths(table_buffer, total_w_twips)

        # Set <w:tblGrid>
        tblGrid = t._tbl.tblGrid
        if tblGrid is not None:
            for cg in list(tblGrid):
                tblGrid.remove(cg)
            for w_tw in col_widths_twips:
                tblGrid.append(parse_xml(f'<w:gridCol {NS_W} w:w="{w_tw}"/>'))

        for r_idx, row_data in enumerate(table_buffer):
            row = t.rows[r_idx]
            trPr = row._tr.get_or_add_trPr()
            trPr.append(parse_xml(f'<w:cantSplit {NS_W}/>'))

            if r_idx == 0 and len(table_buffer) > 1:
                trPr.append(parse_xml(f'<w:tblHeader {NS_W}/>'))

            for c_idx, val in enumerate(row_data):
                if c_idx < len(row.cells):
                    cell = row.cells[c_idx]
                    if c_idx < len(col_widths_twips):
                        cell.width = Inches(col_widths_twips[c_idx] / 1440.0)

                    tcPr = cell._tc.get_or_add_tcPr()

                    # ECMA-376 Schema sequence: <w:shd> MUST precede <w:vAlign>
                    if r_idx == 0:
                        shading = parse_xml(f'<w:shd {NS_W} w:fill="{header_bg_hex}"/>')
                        tcPr.append(shading)

                    if tcPr.find(qn("w:vAlign")) is None:
                        tcPr.append(parse_xml(f'<w:vAlign {NS_W} w:val="center"/>'))

                    cell.text = ""
                    # Handle multiple lines or list items inside table cells
                    cell_lines = re.split(r"<br\s*/?>|\n", str(val))
                    cell_lines = [cl for cl in cell_lines if cl.strip()]

                    if not cell_lines:
                        cp = cell.paragraphs[0] if cell.paragraphs else cell.add_paragraph()
                        cp.paragraph_format.space_before = Pt(2.0)
                        cp.paragraph_format.space_after = Pt(2.0)
                    else:
                        cell_list_tracker = ListContextTracker(doc)
                        for l_idx, c_line in enumerate(cell_lines):
                            cp = cell.paragraphs[0] if (l_idx == 0 and cell.paragraphs) else cell.add_paragraph()
                            c_item = cell_list_tracker.parse_markdown_line(c_line)
                            if c_item:
                                _add_formatted_runs(cp, c_item.text)
                                num_id = cell_list_tracker.get_num_id_for_item(c_item)
                                apply_list_formatting(cp, c_item.item_type, c_item.level, num_id, c_item.is_checked)
                            else:
                                cell_list_tracker.reset_all_lists()
                                _add_formatted_runs(cp, c_line)
                                cp.paragraph_format.space_before = Pt(2.0)
                                cp.paragraph_format.space_after = Pt(2.0)
                                cp.paragraph_format.line_spacing = 1.15

                            if c_idx < len(col_alignments):
                                cp.alignment = col_alignments[c_idx]
                            else:
                                cp.alignment = WD_ALIGN_PARAGRAPH.CENTER if c_idx == 0 else WD_ALIGN_PARAGRAPH.LEFT
                            for cr in cp.runs:
                                cr.font.name = font_family
                                cr.font.size = Pt(10)
                                if r_idx == 0:
                                    cr.bold = True

        table_buffer.clear()
        col_alignments.clear()
        in_table = False

    for idx, line in enumerate(lines):
        if progress_callback and total_lines > 0:
            percent = (idx / total_lines) * 90.0
            progress_callback(percent, f"Đang kết xuất DOCX ({idx + 1}/{total_lines})...")

        s_line = line.strip()

        # Handle Fenced Code Blocks
        if s_line.startswith("```"):
            if in_table:
                flush_table()
            if not in_code_block:
                in_code_block = True
                code_block_lang = s_line.lstrip("`").strip().lower()
                code_block_lines = []
            else:
                in_code_block = False
                _flush_fenced_code_block(doc, code_block_lang, code_block_lines, assets_dir, font_family)
                code_block_lang = ""
                code_block_lines = []
            continue

        if in_code_block:
            code_block_lines.append(line.rstrip("\r\n"))
            continue

        # Handle Tables (only when not in code block and has meaningful cells)
        if s_line.startswith("|") and s_line.endswith("|") and len(s_line) > 1 and "|" in s_line[1:-1]:
            parts = [p.strip() for p in s_line.split("|")[1:-1]]
            if not in_table:
                in_table = True
                table_buffer.append(parts)
            else:
                if all(re.match(r"^[-:]+$", p) for p in parts):
                    col_alignments.clear()
                    for p_al in parts:
                        if p_al.startswith(":") and p_al.endswith(":"):
                            col_alignments.append(WD_ALIGN_PARAGRAPH.CENTER)
                        elif p_al.endswith(":"):
                            col_alignments.append(WD_ALIGN_PARAGRAPH.RIGHT)
                        else:
                            col_alignments.append(WD_ALIGN_PARAGRAPH.LEFT)
                    continue
                table_buffer.append(parts)
            continue
        elif in_table:
            flush_table()

        if not s_line:
            continue

        # Handle Images & Diagrams
        img_match = re.search(r"!\[(.*?)\]\((.*?)\)", s_line)
        if img_match:
            list_tracker.reset_all_lists()
            caption_text = img_match.group(1).strip()
            img_rel_path = img_match.group(2).strip()
            full_img_path = os.path.join(base_dir, img_rel_path)

            try:
                from core.diagrams.bridge import render_diagram_file_to_png, calculate_aspect_dimensions
            except (ImportError, ValueError):
                try:
                    from diagram_bridge import render_diagram_file_to_png, calculate_aspect_dimensions
                except (ImportError, ValueError):
                    render_diagram_file_to_png = None
                    calculate_aspect_dimensions = None

            diag_res = render_diagram_file_to_png(full_img_path, base_dir=base_dir) if render_diagram_file_to_png else None
            if diag_res and diag_res.get("png_path") and os.path.isfile(diag_res["png_path"]):
                png_path = diag_res["png_path"]
                w_px, h_px = diag_res["dimensions_px"]
                w_cm, h_cm = calculate_aspect_dimensions(w_px, h_px, max_w_cm=14.0, max_h_cm=18.0) if calculate_aspect_dimensions else (14.0, 10.0)

                p_img = doc.add_paragraph()
                p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p_img.paragraph_format.space_before = Pt(8)
                p_img.paragraph_format.space_after = Pt(2)
                p_img.add_run().add_picture(png_path, width=Inches(w_cm / 2.54))

                final_cap = caption_text or (diag_res.get("title") if diag_res.get("engine") != "image" else "")
                if final_cap:
                    p_cap = doc.add_paragraph()
                    p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    r_cap = p_cap.add_run(final_cap)
                    r_cap.font.name = font_family
                    r_cap.font.size = Pt(9.5)
                    r_cap.italic = True
                    p_cap.paragraph_format.space_before = Pt(2)
                    p_cap.paragraph_format.space_after = Pt(8)
                continue
            elif os.path.isfile(full_img_path):
                # Direct fallback for standard image file
                p_img = doc.add_paragraph()
                p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p_img.paragraph_format.space_before = Pt(8)
                p_img.paragraph_format.space_after = Pt(2)
                p_img.add_run().add_picture(full_img_path, width=Inches(5.5))
                if caption_text:
                    p_cap = doc.add_paragraph()
                    p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    r_cap = p_cap.add_run(caption_text)
                    r_cap.font.name = font_family
                    r_cap.font.size = Pt(9.5)
                    r_cap.italic = True
                    p_cap.paragraph_format.space_before = Pt(2)
                    p_cap.paragraph_format.space_after = Pt(8)
                continue

        # Table & Figure Captions
        if s_line.startswith("*Table ") or s_line.startswith("*Figure "):
            list_tracker.reset_all_lists()
            caption_text = s_line.strip("*")
            p_cap = doc.add_paragraph()
            p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r_cap = p_cap.add_run(caption_text)
            r_cap.font.name = font_family
            r_cap.font.size = Pt(10)
            r_cap.italic = True
            p_cap.paragraph_format.space_before = Pt(3)
            p_cap.paragraph_format.space_after = Pt(8)
            continue

        # TOC lines with dot leaders
        m_toc = TOC_LINE_REGEX.match(s_line)
        if m_toc:
            list_tracker.reset_all_lists()
            title_p = m_toc.group(1).strip()
            page_p = m_toc.group(2).strip()
            p_toc = doc.add_paragraph()
            _apply_tab_stop_to_paragraph(p_toc, title_p, page_p, font_name=font_family, tab_pos_twips=tab_pos_twips)
            continue

        # Horizontal Rule (Divider)
        if s_line == "---":
            list_tracker.reset_all_lists()
            continue

        # Headings
        if s_line.startswith("# "):
            list_tracker.reset_all_lists()
            h = doc.add_heading(s_line[2:].strip(), level=1)
            h.paragraph_format.space_before = Pt(14)
            h.paragraph_format.space_after = Pt(6)
            h.paragraph_format.keep_with_next = True
        elif s_line.startswith("## "):
            list_tracker.reset_all_lists()
            h = doc.add_heading(s_line[3:].strip(), level=2)
            h.paragraph_format.space_before = Pt(12)
            h.paragraph_format.space_after = Pt(4)
            h.paragraph_format.keep_with_next = True
        elif s_line.startswith("### "):
            list_tracker.reset_all_lists()
            h = doc.add_heading(s_line[4:].strip(), level=3)
            h.paragraph_format.space_before = Pt(8)
            h.paragraph_format.space_after = Pt(3)
            h.paragraph_format.keep_with_next = True
        elif s_line.startswith("#### "):
            list_tracker.reset_all_lists()
            h = doc.add_heading(s_line[5:].strip(), level=4)
            h.paragraph_format.space_before = Pt(6)
            h.paragraph_format.space_after = Pt(2)
            h.paragraph_format.keep_with_next = True
        elif s_line.startswith("##### "):
            list_tracker.reset_all_lists()
            h_text = s_line[6:].strip()
            try:
                h = doc.add_heading(h_text, level=5)
            except Exception:
                h = doc.add_heading(h_text, level=4)
            h.paragraph_format.space_before = Pt(5)
            h.paragraph_format.space_after = Pt(2)
            h.paragraph_format.keep_with_next = True
            for r in h.runs:
                r.bold = True
                r.font.name = font_family
                r.font.color.rgb = h_color
        elif s_line.startswith("###### "):
            list_tracker.reset_all_lists()
            h_text = s_line[7:].strip()
            try:
                h = doc.add_heading(h_text, level=6)
            except Exception:
                h = doc.add_heading(h_text, level=4)
            h.paragraph_format.space_before = Pt(4)
            h.paragraph_format.space_after = Pt(2)
            h.paragraph_format.keep_with_next = True
            for r in h.runs:
                r.bold = True
                r.font.name = font_family
                r.font.color.rgb = h_color
        else:
            # Check for list items (Ordered, Bullet, Checkbox with multi-level indentation)
            list_item = list_tracker.parse_markdown_line(line)
            if list_item:
                p = doc.add_paragraph()
                _add_formatted_runs(p, list_item.text, font_name=font_family)
                num_id = list_tracker.get_num_id_for_item(list_item)
                apply_list_formatting(p, list_item.item_type, list_item.level, num_id, list_item.is_checked)
                if len(s_line) > 75:
                    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                continue

            # Plain paragraph: reset list tracker (Zero auto-numbering on plain text)
            list_tracker.reset_all_lists()
            p = doc.add_paragraph()
            _add_formatted_runs(p, s_line, font_name=font_family)
            p.paragraph_format.line_spacing = 1.15
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(4)
            if s_line.startswith("$$") and s_line.endswith("$$"):
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p.paragraph_format.space_before = Pt(6)
                p.paragraph_format.space_after = Pt(6)

            # Detect Signature block at the end
            if "NGƯỜI LẬP KẾ HOẠCH" in s_line:
                in_signature_block = True
                p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                p.paragraph_format.space_before = Pt(14)
            elif in_signature_block:
                p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                if "(Đã ký)" in s_line:
                    p.paragraph_format.space_after = Pt(18)  # Leave room for signature
            elif s_line.startswith("**Kính gửi:") or s_line.startswith("**Ứng viên:") or s_line.startswith("**Vị trí") or s_line.startswith("**Tài liệu") or s_line.startswith("**Thời điểm") or s_line.startswith("**Email:"):
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            elif len(s_line) > 60:
                p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            else:
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT

    if in_table:
        flush_table()

    if in_code_block and code_block_lines:
        _flush_fenced_code_block(doc, code_block_lang, code_block_lines, assets_dir, font_family)

    doc.save(docx_path)
    if progress_callback:
        progress_callback(100.0, f"Đã tạo file Word: {os.path.basename(docx_path)}")

    return True, docx_path


def _apply_tab_stop_to_paragraph(
    p: docx.text.paragraph.Paragraph,
    title_text: str,
    page_num_text: str,
    font_name: str = "Calibri",
    font_size_pt: float = 11.0,
    tab_pos_twips: int = 9360
) -> None:
    """Format paragraph with right-aligned Tab Stop with dot leader."""
    p.text = ""
    r_title = p.add_run(title_text)
    r_title.font.name = font_name
    r_title.font.size = Pt(font_size_pt)

    r_tab = p.add_run("\t")
    r_tab.font.name = font_name
    r_tab.font.size = Pt(font_size_pt)

    r_num = p.add_run(page_num_text)
    r_num.font.name = font_name
    r_num.font.size = Pt(font_size_pt)

    pPr = p._p.get_or_add_pPr()
    pPr.append(parse_xml(
        f'<w:tabs {NS_W}>'
        f'<w:tab w:val="right" w:leader="dot" w:pos="{tab_pos_twips}"/>'
        f'</w:tabs>'
    ))
    pPr.append(parse_xml(f'<w:spacing {NS_W} w:before="60" w:after="40" w:line="240" w:lineRule="auto"/>'))


SYMBOL_FALLBACK_REGEX = re.compile(
    r"([\u2190-\u21FF\u25A0-\u25FF\u2600-\u27BF]|\U0001F300-\U0001FAFF)"
)


def _add_symbol_fallback_runs(
    p: docx.text.paragraph.Paragraph,
    text: str,
    base_font: Optional[str] = None,
    bold: bool = False,
    italic: bool = False,
    color_rgb: Optional[RGBColor] = None,
    size_pt: Optional[float] = None,
) -> None:
    """
    Appends text runs with fallback to Segoe UI Symbol / Segoe UI Emoji (Windows)
    or DejaVu Sans / Noto Color Emoji (Linux) for arrows and emojis.
    """
    if not text:
        return
    is_win = sys.platform == "win32"
    tokens = SYMBOL_FALLBACK_REGEX.split(text)
    for tok in tokens:
        if not tok:
            continue
        r = p.add_run(tok)
        r.bold = bold
        r.italic = italic
        if size_pt:
            r.font.size = Pt(size_pt)
        if color_rgb:
            r.font.color.rgb = color_rgb

        if SYMBOL_FALLBACK_REGEX.match(tok):
            if tok in "↔→←↑↓↻⇄⇒⇌":
                target_font = "Segoe UI Symbol" if is_win else "DejaVu Sans"
            else:
                target_font = "Segoe UI Emoji" if is_win else "Noto Color Emoji"
            r.font.name = target_font
            try:
                rPr = r._r.get_or_add_rPr()
                rFonts = parse_xml(f'<w:rFonts {NS_W} w:ascii="{target_font}" w:hAnsi="{target_font}" w:cs="{target_font}"/>')
                rPr.append(rFonts)
            except Exception:
                pass
        elif base_font:
            r.font.name = base_font


def _add_formatted_runs(p: docx.text.paragraph.Paragraph, text: str, font_name: Optional[str] = None) -> None:
    """Parse inline Markdown links, line breaks, backticks, bold, italic, and LaTeX math into docx runs with font fallback."""
    cleaned_text = re.sub(r"\[(.*?)\]\(.*?\)", r"\1", text)
    cleaned_text = re.sub(r"<br\s*/?>", "\n", cleaned_text)
    cleaned_text = re.sub(r"`(.*?)`", r"\1", cleaned_text)
    cleaned_text = cleaned_text.replace("`", "")

    if "$" in cleaned_text:
        add_math_to_paragraph(p, cleaned_text, font_name=font_name or "Calibri")
        return

    tokens = re.split(r"(\*\*.*?\*\*|\*.*?\*)", cleaned_text)
    for token in tokens:
        if not token:
            continue
        if token.startswith("**") and token.endswith("**") and len(token) >= 4:
            _add_symbol_fallback_runs(p, token[2:-2], base_font=font_name, bold=True)
        elif token.startswith("*") and token.endswith("*") and len(token) >= 2:
            _add_symbol_fallback_runs(p, token[1:-1], base_font=font_name, italic=True)
        else:
            _add_symbol_fallback_runs(p, token, base_font=font_name)


# ==============================================================================
# 3. Facade wrappers (Compatible with CLI & GUI)
# ==============================================================================
def docx_to_markdown(
    docx_path: str,
    md_path: Optional[str] = None,
    progress_callback: Optional[Callable[[float, str], None]] = None
) -> Tuple[bool, str]:
    """Wraps docx_to_decoupled: outputs both [name].md and [name].style.yaml."""
    ok, out_md, _ = docx_to_decoupled(docx_path, md_path=md_path, progress_callback=progress_callback)
    return ok, out_md


def markdown_to_docx(
    md_path: str,
    docx_path: Optional[str] = None,
    style_yaml_path: Optional[str] = None,
    progress_callback: Optional[Callable[[float, str], None]] = None
) -> Tuple[bool, str]:
    """Wraps decoupled_to_docx: auto-discovers [name].style.yaml."""
    return decoupled_to_docx(md_path, docx_path=docx_path, style_yaml_path=style_yaml_path, progress_callback=progress_callback)


def pdf_to_markdown(
    pdf_path: str,
    md_path: Optional[str] = None,
    progress_callback: Optional[Callable[[float, str], None]] = None
) -> Tuple[bool, str]:
    """
    Full 3-stage PDF -> Markdown pipeline:
    1. Converts PDF -> pristine DOCX using Flow Engine + Smart Post-Processor v6.
    2. Bóc tách DOCX -> [name].md (Data) + [name].style.yaml (Style).
    """
    if not os.path.isfile(pdf_path):
        return False, f"File không tồn tại: {pdf_path}"

    if not md_path:
        md_path = os.path.splitext(pdf_path)[0] + ".md"

    # Temporary pristine DOCX
    temp_docx = os.path.splitext(md_path)[0] + "_temp_stage1.docx"

    try:
        from .converter_engine import convert_pdf_to_docx_flow
    except ImportError:
        from converter_engine import convert_pdf_to_docx_flow

    if progress_callback:
        progress_callback(10.0, "Tầng 1/2: Chuyển đổi PDF sang DOCX chuẩn 100% (Post-Processor v6)...")

    ok_d, res_d = convert_pdf_to_docx_flow(pdf_path, temp_docx, progress_callback=progress_callback)
    if not ok_d:
        return False, f"Lỗi tại Tầng 1 (PDF -> DOCX): {res_d}"

    if progress_callback:
        progress_callback(60.0, "Tầng 2/2: Bóc tách DOCX sang Data .md và Style .style.yaml...")

    style_yaml_path = os.path.splitext(md_path)[0] + ".style.yaml"
    ok_md, out_md, out_yaml = docx_to_decoupled(temp_docx, md_path, style_yaml_path, progress_callback=progress_callback)

    if os.path.exists(temp_docx):
        try:
            os.remove(temp_docx)
        except Exception:
            pass

    return ok_md, out_md


# ==============================================================================
# 4. DOCX -> PDF & Markdown -> PDF
# ==============================================================================
def _find_edge_path() -> Optional[str]:
    import shutil
    candidates = [
        r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
        r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
        shutil.which("msedge"),
        shutil.which("msedge.exe"),
    ]
    for c in candidates:
        if c and os.path.isfile(c):
            return c
    return None


def _md_to_html_printable(md_text: str, base_dir: str) -> str:
    """Converts GFM markdown to self-contained HTML for PDF printing."""
    import html as html_lib
    html_lines = [
        "<!DOCTYPE html>",
        "<html lang=\"vi\">",
        "<head>",
        "  <meta charset=\"utf-8\">",
        "  <title>Document</title>",
        "  <style>",
        "    @page {",
        "      size: A4 portrait;",
        "      margin: 20mm 15mm 20mm 20mm;",
        "    }",
        "    @media print {",
        "      body { -webkit-print-color-adjust: exact; print-color-adjust: exact; }",
        "    }",
        "    body {",
        "      font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, 'Helvetica Neue', Arial, sans-serif;",
        "      font-size: 11pt;",
        "      line-height: 1.6;",
        "      color: #0f172a;",
        "      margin: 0;",
        "      padding: 0;",
        "    }",
        "    h1 {",
        "      font-size: 20pt;",
        "      font-weight: 700;",
        "      color: #0051e2;",
        "      border-bottom: 2px solid #0051e2;",
        "      padding-bottom: 6px;",
        "      margin-top: 24px;",
        "      margin-bottom: 12px;",
        "      page-break-after: avoid;",
        "    }",
        "    h2 {",
        "      font-size: 15pt;",
        "      font-weight: 600;",
        "      color: #1e3a8a;",
        "      border-bottom: 1px solid #cbd5e1;",
        "      padding-bottom: 4px;",
        "      margin-top: 20px;",
        "      margin-bottom: 10px;",
        "      page-break-after: avoid;",
        "    }",
        "    h3 {",
        "      font-size: 12.5pt;",
        "      font-weight: 600;",
        "      color: #334155;",
        "      margin-top: 16px;",
        "      margin-bottom: 8px;",
        "      page-break-after: avoid;",
        "    }",
        "    h4, h5, h6 {",
        "      font-size: 11.5pt;",
        "      font-weight: 600;",
        "      color: #475569;",
        "      margin-top: 14px;",
        "      margin-bottom: 6px;",
        "      page-break-after: avoid;",
        "    }",
        "    p {",
        "      margin: 8px 0;",
        "      text-align: justify;",
        "    }",
        "    ul, ol {",
        "      margin: 8px 0 8px 24px;",
        "      padding-left: 0;",
        "    }",
        "    li {",
        "      margin-bottom: 4px;",
        "    }",
        "    table {",
        "      width: 100%;",
        "      border-collapse: collapse;",
        "      margin: 16px 0;",
        "      page-break-inside: avoid;",
        "      font-size: 10pt;",
        "    }",
        "    th, td {",
        "      border: 1px solid #cbd5e1;",
        "      padding: 8px 10px;",
        "      text-align: left;",
        "      vertical-align: top;",
        "    }",
        "    th {",
        "      background-color: #f1f5f9;",
        "      color: #0f172a;",
        "      font-weight: 600;",
        "    }",
        "    tr:nth-child(even) td {",
        "      background-color: #f8fafc;",
        "    }",
        "    code {",
        "      font-family: Consolas, 'Courier New', monospace;",
        "      font-size: 9.5pt;",
        "      background: #f1f5f9;",
        "      padding: 2px 4px;",
        "      border-radius: 3px;",
        "      color: #0f172a;",
        "    }",
        "    pre {",
        "      background: #0f172a;",
        "      color: #f8fafc;",
        "      padding: 12px 16px;",
        "      border-radius: 6px;",
        "      overflow-x: auto;",
        "      font-size: 9pt;",
        "      page-break-inside: avoid;",
        "      margin: 14px 0;",
        "    }",
        "    pre code {",
        "      background: transparent;",
        "      color: #f8fafc;",
        "      padding: 0;",
        "      border-radius: 0;",
        "    }",
        "    blockquote {",
        "      border-left: 4px solid #0051e2;",
        "      margin: 12px 0;",
        "      padding: 6px 14px;",
        "      background: #f8fafc;",
        "      color: #475569;",
        "    }",
        "    hr {",
        "      border: 0;",
        "      border-top: 1px solid #cbd5e1;",
        "      margin: 20px 0;",
        "    }",
        "    img {",
        "      max-width: 100%;",
        "      height: auto;",
        "      display: block;",
        "      margin: 16px auto;",
        "      page-break-inside: avoid;",
        "      border-radius: 4px;",
        "    }",
        "  </style>",
        "</head>",
        "<body>"
    ]

    in_code_block = False
    code_block_lang = ""
    code_block_lines: list[str] = []
    in_table = False
    in_ul = False
    in_ol = False

    def close_lists() -> list[str]:
        nonlocal in_ul, in_ol
        res = []
        if in_ul:
            res.append("</ul>")
            in_ul = False
        if in_ol:
            res.append("</ol>")
            in_ol = False
        return res

    lines = md_text.splitlines()
    i = 0
    while i < len(lines):
        line = lines[i]

        if line.startswith("```"):
            html_lines.extend(close_lists())
            if in_table:
                html_lines.append("</tbody></table>")
                in_table = False
            if in_code_block:
                code_content = "\n".join(code_block_lines).strip()
                code_lang = (code_block_lang or "").strip().lower()
                diag_res = None
                if code_lang in (
                    "mermaid", "mmd", "plantuml", "puml", "drawio", "mxgraph", "xml",
                    "canvas", "spec", "diagram-spec", "diagram", "json"
                ):
                    try:
                        from core.diagrams.bridge import render_diagram_code_to_png
                    except (ImportError, ValueError):
                        try:
                            from diagram_bridge import render_diagram_code_to_png
                        except (ImportError, ValueError):
                            render_diagram_code_to_png = None

                    if render_diagram_code_to_png:
                        diag_res = render_diagram_code_to_png(code_content, code_lang, base_dir=base_dir)

                if diag_res and diag_res.get("png_path") and os.path.isfile(diag_res["png_path"]):
                    diag_png = diag_res["png_path"]
                    img_uri = f"file:///{diag_png.replace(os.sep, '/')}"
                    cap_text = diag_res.get("title") or f"Sơ đồ {code_lang.capitalize()}"
                    html_lines.append('<div class="diagram-figure" style="text-align: center; margin: 16px auto;">')
                    html_lines.append(f'  <img src="{img_uri}" alt="{html_lib.escape(cap_text)}" style="max-width: 100%; height: auto; display: block; margin: 0 auto; border-radius: 4px;" />')
                    html_lines.append(f'  <p style="font-size: 9.5pt; font-style: italic; color: #64748b; margin-top: 6px;">{html_lib.escape(cap_text)}</p>')
                    html_lines.append('</div>')
                else:
                    html_lines.append("<pre><code>")
                    for cl in code_block_lines:
                        html_lines.append(html_lib.escape(cl))
                    html_lines.append("</code></pre>")

                in_code_block = False
                code_block_lines = []
                code_block_lang = ""
            else:
                in_code_block = True
                code_block_lang = line[3:].strip()
                code_block_lines = []
            i += 1
            continue

        if in_code_block:
            code_block_lines.append(line)
            i += 1
            continue

        stripped = line.strip()
        if stripped.startswith("|") and stripped.endswith("|"):
            html_lines.extend(close_lists())
            if not in_table:
                in_table = True
                html_lines.append("<table>")
                cells = [c.strip() for c in stripped[1:-1].split("|")]
                th_html = "".join(f"<th>{html_lib.escape(c)}</th>" for c in cells)
                html_lines.append(f"<thead><tr>{th_html}</tr></thead><tbody>")
                if i + 1 < len(lines) and "---" in lines[i+1]:
                    i += 1
            else:
                cells = [c.strip() for c in stripped[1:-1].split("|")]
                td_html = "".join(f"<td>{html_lib.escape(c)}</td>" for c in cells)
                html_lines.append(f"<tr>{td_html}</tr>")
            i += 1
            continue
        elif in_table:
            html_lines.append("</tbody></table>")
            in_table = False

        if not stripped:
            html_lines.extend(close_lists())
            i += 1
            continue

        if stripped.startswith("# "):
            html_lines.extend(close_lists())
            html_lines.append(f"<h1>{html_lib.escape(stripped[2:])}</h1>")
        elif stripped.startswith("## "):
            html_lines.extend(close_lists())
            html_lines.append(f"<h2>{html_lib.escape(stripped[3:])}</h2>")
        elif stripped.startswith("### "):
            html_lines.extend(close_lists())
            html_lines.append(f"<h3>{html_lib.escape(stripped[4:])}</h3>")
        elif stripped.startswith("#### "):
            html_lines.extend(close_lists())
            html_lines.append(f"<h4>{html_lib.escape(stripped[5:])}</h4>")
        elif stripped.startswith("##### "):
            html_lines.extend(close_lists())
            html_lines.append(f"<h5>{html_lib.escape(stripped[6:])}</h5>")
        elif stripped.startswith("- ") or stripped.startswith("* "):
            if in_ol:
                html_lines.append("</ol>")
                in_ol = False
            if not in_ul:
                html_lines.append("<ul>")
                in_ul = True
            content = stripped[2:]
            content = re.sub(r"\*\*([^*]+)\*\*", r"<strong>\1</strong>", content)
            content = re.sub(r"`([^`]+)`", r"<code>\1</code>", content)
            html_lines.append(f"<li>{content}</li>")
        elif re.match(r"^\d+\.\s+", stripped):
            if in_ul:
                html_lines.append("</ul>")
                in_ul = False
            if not in_ol:
                html_lines.append("<ol>")
                in_ol = True
            content = re.sub(r"^\d+\.\s+", "", stripped)
            content = re.sub(r"\*\*([^*]+)\*\*", r"<strong>\1</strong>", content)
            content = re.sub(r"`([^`]+)`", r"<code>\1</code>", content)
            html_lines.append(f"<li>{content}</li>")
        elif stripped.startswith("> "):
            html_lines.extend(close_lists())
            content = stripped[2:]
            content = re.sub(r"\*\*([^*]+)\*\*", r"<strong>\1</strong>", content)
            html_lines.append(f"<blockquote>{content}</blockquote>")
        elif stripped in ("---", "***"):
            html_lines.extend(close_lists())
            html_lines.append("<hr>")
        elif re.match(r"^!\[(.*?)\]\((.*?)\)$", stripped):
            html_lines.extend(close_lists())
            m = re.match(r"^!\[(.*?)\]\((.*?)\)$", stripped)
            if m:
                alt, img_url = m.group(1), m.group(2)
                try:
                    from core.diagrams.bridge import render_diagram_file_to_png
                except (ImportError, ValueError):
                    try:
                        from diagram_bridge import render_diagram_file_to_png
                    except (ImportError, ValueError):
                        render_diagram_file_to_png = None

                diag_file_res = render_diagram_file_to_png(img_url, base_dir=base_dir) if render_diagram_file_to_png else None
                if diag_file_res and diag_file_res.get("png_path") and os.path.isfile(diag_file_res["png_path"]):
                    resolved_img = diag_file_res["png_path"]
                    alt = alt or (diag_file_res.get("title") if diag_file_res.get("engine") != "image" else "")
                else:
                    resolved_img = img_url if os.path.isabs(img_url) else os.path.normpath(os.path.join(base_dir, img_url))

                img_uri = f"file:///{resolved_img.replace(os.sep, '/')}"
                html_lines.append('<div class="diagram-figure" style="text-align: center; margin: 16px auto;">')
                html_lines.append(f'  <img src="{img_uri}" alt="{html_lib.escape(alt)}" style="max-width: 100%; height: auto; display: block; margin: 0 auto; border-radius: 4px;" />')
                if alt:
                    html_lines.append(f'  <p style="font-size: 9.5pt; font-style: italic; color: #64748b; margin-top: 6px;">{html_lib.escape(alt)}</p>')
                html_lines.append('</div>')

        else:
            html_lines.extend(close_lists())
            text = stripped
            text = re.sub(r"\*\*([^*]+)\*\*", r"<strong>\1</strong>", text)
            text = re.sub(r"`([^`]+)`", r"<code>\1</code>", text)
            text = re.sub(r"\*([^*]+)\*", r"<em>\1</em>", text)
            html_lines.append(f"<p>{text}</p>")

        i += 1

    html_lines.extend(close_lists())
    if in_table:
        html_lines.append("</tbody></table>")
    if in_code_block:
        html_lines.append("</code></pre>")

    html_lines.append("</body></html>")
    return "\n".join(html_lines)


def _docx_to_pdf_via_edge(
    docx_path: str,
    pdf_path: str,
    progress_callback: Optional[Callable[[float, str], None]] = None
) -> Tuple[bool, str]:
    import subprocess
    import shutil

    edge_exe = _find_edge_path()
    if not edge_exe:
        return False, "Không tìm thấy Microsoft Edge trên hệ thống để xuất bản PDF."

    abs_docx = os.path.abspath(docx_path)
    abs_pdf = os.path.abspath(pdf_path)
    temp_dir = os.path.join(os.path.dirname(abs_pdf), f".temp_pdf_{os.getpid()}")
    os.makedirs(temp_dir, exist_ok=True)

    temp_md = os.path.join(temp_dir, "doc.md")
    temp_yaml = os.path.join(temp_dir, "doc.style.yaml")
    temp_html = os.path.join(temp_dir, "doc.html")

    try:
        if progress_callback:
            progress_callback(25.0, "Đang bóc tách nội dung DOCX sang Markdown...")

        ok_md, out_md, _ = docx_to_decoupled(abs_docx, temp_md, temp_yaml)
        if not ok_md or not os.path.isfile(temp_md):
            return False, f"Không thể giải mã DOCX: {out_md}"

        with open(temp_md, "r", encoding="utf-8") as f:
            md_text = f.read()

        if progress_callback:
            progress_callback(50.0, "Đang định dạng trang in A4...")

        html_content = _md_to_html_printable(md_text, base_dir=temp_dir)
        with open(temp_html, "w", encoding="utf-8") as f:
            f.write(html_content)

        if progress_callback:
            progress_callback(75.0, "Đang xuất bản PDF chuẩn vector qua Edge Engine...")

        cmd = [
            edge_exe,
            "--headless",
            "--disable-gpu",
            "--no-pdf-header-footer",
            f"--print-to-pdf={abs_pdf}",
            temp_html
        ]

        proc = subprocess.run(cmd, capture_output=True, text=True, timeout=30)
        if proc.returncode != 0 and not os.path.isfile(abs_pdf):
            return False, f"Lỗi Edge Print to PDF (code {proc.returncode}): {proc.stderr}"

        if not os.path.isfile(abs_pdf) or os.path.getsize(abs_pdf) == 0:
            return False, "Tệp PDF xuất ra bị rỗng hoặc không tồn tại."

        if progress_callback:
            progress_callback(100.0, f"Hoàn tất xuất PDF: {os.path.basename(abs_pdf)}")

        return True, abs_pdf

    except Exception as exc:
        return False, f"Lỗi trong quá trình xuất PDF: {exc}"
    finally:
        try:
            shutil.rmtree(temp_dir, ignore_errors=True)
        except Exception:
            pass


def docx_to_pdf(
    docx_path: str,
    pdf_path: Optional[str] = None,
    progress_callback: Optional[Callable[[float, str], None]] = None
) -> Tuple[bool, str]:
    if not os.path.isfile(docx_path):
        return False, f"File không tồn tại: {docx_path}"

    if not pdf_path:
        pdf_path = os.path.splitext(docx_path)[0] + ".pdf"

    # Tier 1: Microsoft Word COM Automation (Pixel-perfect 100% vector fidelity on Windows)
    if sys.platform == "win32":
        try:
            if progress_callback:
                progress_callback(50.0, "Đang xuất bản PDF qua Microsoft Word COM...")
            import pythoncom
            pythoncom.CoInitialize()
            import win32com.client as win32
            word = win32.Dispatch("Word.Application")
            word.Visible = False
            word.DisplayAlerts = 0

            abs_docx = os.path.abspath(docx_path)
            abs_pdf = os.path.abspath(pdf_path)

            doc = word.Documents.Open(abs_docx, ReadOnly=True)
            doc.SaveAs2(abs_pdf, FileFormat=17)
            doc.Close(SaveChanges=False)
            word.Quit()
            pythoncom.CoUninitialize()

            if os.path.isfile(abs_pdf) and os.path.getsize(abs_pdf) > 0:
                if progress_callback:
                    progress_callback(100.0, f"Hoàn tất xuất PDF (Word COM): {os.path.basename(abs_pdf)}")
                return True, abs_pdf
        except Exception as exc:
            pass

    # Tier 2: LibreOffice Headless (Cross-platform on Linux/Docker/Windows)
    soffice_candidates = [
        r"C:\Program Files\LibreOffice\program\soffice.com",
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        shutil.which("soffice"),
        shutil.which("libreoffice"),
    ]
    soffice_path = next((c for c in soffice_candidates if c and os.path.isfile(c)), None)
    if soffice_path:
        try:
            if progress_callback:
                progress_callback(50.0, "Chuyển sang LibreOffice export PDF...")
            abs_docx = os.path.abspath(docx_path)
            out_dir = os.path.dirname(os.path.abspath(pdf_path))
            import subprocess
            cmd = [soffice_path, "--headless", "--nologo", "--convert-to", "pdf", "--outdir", out_dir, abs_docx]
            subprocess.run(cmd, capture_output=True, text=True, timeout=30)
            expected_out = os.path.join(out_dir, os.path.splitext(os.path.basename(docx_path))[0] + ".pdf")
            if os.path.isfile(expected_out):
                if expected_out != os.path.abspath(pdf_path):
                    shutil.move(expected_out, os.path.abspath(pdf_path))
                if progress_callback:
                    progress_callback(100.0, f"Hoàn tất xuất PDF (LibreOffice): {os.path.basename(pdf_path)}")
                return True, pdf_path
        except Exception:
            pass

    # Tier 3: Microsoft Edge Headless Print to PDF (Vector fallback)
    edge_exe = _find_edge_path()
    if edge_exe:
        ok_edge, res_edge = _docx_to_pdf_via_edge(docx_path, pdf_path, progress_callback)
        if ok_edge and os.path.isfile(pdf_path) and os.path.getsize(pdf_path) > 0:
            return True, pdf_path

    return False, "Không thể xuất PDF: Cần Microsoft Word, LibreOffice hoặc Microsoft Edge để chuyển đổi DOCX sang PDF."


def markdown_to_pdf(
    md_path: str,
    pdf_path: Optional[str] = None,
    style_yaml_path: Optional[str] = None,
    progress_callback: Optional[Callable[[float, str], None]] = None
) -> Tuple[bool, str]:
    if not pdf_path:
        pdf_path = os.path.splitext(md_path)[0] + ".pdf"

    temp_docx = os.path.splitext(md_path)[0] + "_temp.docx"

    if progress_callback:
        progress_callback(10.0, "Bước 1/2: Biên dịch Markdown + Style.yaml sang DOCX...")

    ok_docx, res_docx = markdown_to_docx(md_path, temp_docx, style_yaml_path=style_yaml_path, progress_callback=progress_callback)
    if not ok_docx:
        return False, res_docx

    if progress_callback:
        progress_callback(60.0, "Bước 2/2: Xuất DOCX sang PDF chất lượng cao...")

    ok_pdf, res_pdf = docx_to_pdf(temp_docx, pdf_path, progress_callback)

    if os.path.exists(temp_docx):
        try:
            os.remove(temp_docx)
        except Exception:
            pass

    return ok_pdf, res_pdf
