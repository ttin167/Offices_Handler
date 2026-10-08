"""
docx_to_pptx.py — Cross-Format Synthesis Engine: Word (.docx) ➔ PowerPoint (.pptx).

Automated conversion of structured Word documents into high-fidelity 16:9 presentation slide decks.
Conforms to Enterprise PowerPoint Invariants:
- PPTX_INV_01: Strict 16:9 Widescreen (13.333 x 7.5 inches)
- PPTX_INV_02: Aspect-Ratio Lock on Technical Diagrams & Images
- PPTX_INV_03: Template-driven Master Slide Inheritance (.potx / .pptx)
- PPTX_INV_04: Zero Text Overflow & Smart Bullet Chunking (max 4-5 per slide)
- PPTX_INV_05: Mathematical Centering for Diagrams & Figures
- PPTX_INV_06: Native OpenXML Tables with Pagination (>7 rows with header repeating)
- PPTX_INV_07: Windows File Lock Resilience
- PPTX_INV_08: Spec-First JSON Supremacy (AST ➔ slide_spec.json ➔ .pptx)
"""

from __future__ import annotations

import json
import os
import re
import sys
import time
from typing import Any, Dict, List, Optional, Tuple

if sys.platform == "win32":
    try:
        if hasattr(sys.stdout, "reconfigure"):
            getattr(sys.stdout, "reconfigure")(encoding="utf-8", errors="replace")
        if hasattr(sys.stderr, "reconfigure"):
            getattr(sys.stderr, "reconfigure")(encoding="utf-8", errors="replace")
    except Exception:
        pass

import docx
from docx import Document
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph
from docx.table import Table

try:
    from .pptx_writer import write_pptx_from_spec, get_theme, THEMES
except (ImportError, ValueError):
    from pptx_writer import write_pptx_from_spec, get_theme, THEMES


# ---------------------------------------------------------------------------
# Heuristic & Regex Matchers
# ---------------------------------------------------------------------------

RE_HEADING_1 = re.compile(
    r"^(Chương\s+[0-9IVXLCDM]+|Chapter\s+[0-9IVXLCDM]+|Phần\s+[0-9IVXLCDM]+|Part\s+[0-9IVXLCDM]+|[0-9]+\.\s+[^\d])",
    re.IGNORECASE
)

RE_FIGURE_CAPTION = re.compile(
    r"^(Hình|Figure|Fig\.|Ảnh|Sơ đồ|Biểu đồ|Chart)\s*([0-9]+[\.\-]?[0-9]*)\s*[:\.\-]?\s*(.*)",
    re.IGNORECASE
)

RE_TABLE_CAPTION = re.compile(
    r"^(Bảng|Table|Tab\.)\s*([0-9]+[\.\-]?[0-9]*)\s*[:\.\-]?\s*(.*)",
    re.IGNORECASE
)


# ---------------------------------------------------------------------------
# Extraction Helpers
# ---------------------------------------------------------------------------

def _extract_embedded_images(doc: Document, output_dir: str) -> Dict[str, str]:
    """
    Extracts all embedded media parts from the docx package and saves them to disk.
    Returns mapping from r_id -> extracted image file path.
    """
    os.makedirs(output_dir, exist_ok=True)
    r_id_to_path: Dict[str, str] = {}

    for rel in doc.part.rels.values():
        if "image" in rel.target_ref.lower():
            try:
                target_blob = rel.target_part.blob
                ext = os.path.splitext(rel.target_ref)[1] or ".png"
                img_filename = f"img_{rel.rId}{ext}"
                out_path = os.path.join(output_dir, img_filename)
                with open(out_path, "wb") as f:
                    f.write(target_blob)
                r_id_to_path[rel.rId] = os.path.abspath(out_path)
            except Exception:
                continue

    return r_id_to_path


def _is_heading(para: Paragraph) -> Tuple[bool, int]:
    """
    Determines if a paragraph is a heading and returns (is_heading, level).
    Level 1 = Major Chapter/Section, Level 2 = Subsection, Level 3 = Sub-subsection.
    """
    style_name = para.style.name.lower() if para.style else ""
    text = para.text.strip()

    if not text:
        return False, 0

    if "heading 1" in style_name or "tiêu đề 1" in style_name:
        return True, 1
    if "heading 2" in style_name or "tiêu đề 2" in style_name:
        return True, 2
    if "heading 3" in style_name or "tiêu đề 3" in style_name:
        return True, 3
    if "title" in style_name:
        return True, 1

    # Fallback heuristic: check numbering pattern
    if RE_HEADING_1.match(text) and len(text) < 120:
        return True, 1

    # Pattern like "1.1 Subsection" or "2.3. Details"
    if re.match(r"^[0-9]+\.[0-9]+\s+[^\d]", text) and len(text) < 120:
        return True, 2
    if re.match(r"^[0-9]+\.[0-9]+\.[0-9]+\s+[^\d]", text) and len(text) < 120:
        return True, 3

    return False, 0


def _split_into_sentences(text: str) -> List[str]:
    """Splits a paragraph into clean sentences for bullet points."""
    # Split by period, exclamation, or question mark followed by space or end
    raw = re.split(r"(?<=[.!?])\s+", text.strip())
    sentences = [s.strip() for s in raw if len(s.strip()) > 10]
    return sentences if sentences else [text.strip()]


# ---------------------------------------------------------------------------
# Document Tree Parsing
# ---------------------------------------------------------------------------

def extract_docx_structure(docx_path: str, asset_dir: Optional[str] = None) -> Dict[str, Any]:
    """
    Parses a Word (.docx) document into a hierarchical chapter-section tree
    ready for slide deck synthesis.
    """
    if not os.path.isfile(docx_path):
        raise FileNotFoundError(f"Document not found: {docx_path}")

    doc = Document(docx_path)
    base_name = os.path.splitext(os.path.basename(docx_path))[0]

    if not asset_dir:
        asset_dir = os.path.join(os.path.dirname(os.path.abspath(docx_path)), f".{base_name}_assets")

    image_map = _extract_embedded_images(doc, asset_dir)

    # Document Metadata extraction
    doc_props = doc.core_properties
    doc_title = doc_props.title or ""
    author = doc_props.author or ""

    chapters: List[Dict[str, Any]] = []
    current_chapter: Dict[str, Any] = {
        "title": "Giới thiệu / Tổng quan",
        "subsections": []
    }
    current_section: Dict[str, Any] = {
        "title": "Tổng quan",
        "elements": []
    }

    body = doc.element.body
    for child in body:
        tag = child.tag

        if tag == qn("w:p"):
            p = Paragraph(child, doc)
            text = p.text.strip()

            # Check drawings / images
            drawings = child.findall(f".//{qn('w:drawing')}")
            if drawings:
                for drawing in drawings:
                    blip = drawing.find(f".//{{{qn('a:blip')}}}")
                    if blip is not None:
                        r_id = blip.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed")
                        if r_id and r_id in image_map:
                            current_section["elements"].append({
                                "type": "image",
                                "path": image_map[r_id],
                                "caption": ""
                            })
                continue

            if not text:
                continue

            # First title discovery if metadata was empty
            if not doc_title and len(text) < 120 and ("title" in p.style.name.lower() or _is_heading(p)[1] == 1):
                doc_title = text
                continue

            # Check caption for previous image or table
            cap_match = RE_FIGURE_CAPTION.match(text)
            if cap_match and current_section["elements"] and current_section["elements"][-1]["type"] == "image":
                current_section["elements"][-1]["caption"] = text
                continue

            is_hd, level = _is_heading(p)
            if is_hd:
                if level == 1:
                    # Finalize current section and chapter
                    if current_section["elements"]:
                        current_chapter["subsections"].append(current_section)
                    if current_chapter["subsections"]:
                        chapters.append(current_chapter)

                    current_chapter = {
                        "title": text,
                        "subsections": []
                    }
                    current_section = {
                        "title": text,
                        "elements": []
                    }
                elif level in (2, 3):
                    if current_section["elements"]:
                        current_chapter["subsections"].append(current_section)
                    current_section = {
                        "title": text,
                        "elements": []
                    }
            else:
                # Normal paragraph or bullet point
                is_bullet = "list" in p.style.name.lower() or text.startswith(("•", "-", "*", "+"))
                clean_text = re.sub(r"^[•\-\*\+]\s*", "", text).strip()
                current_section["elements"].append({
                    "type": "bullet" if is_bullet else "paragraph",
                    "text": clean_text
                })

        elif tag == qn("w:tbl"):
            tbl = Table(child, doc)
            rows_data: List[List[str]] = []
            for r in tbl.rows:
                row_vals = [cell.text.strip().replace("\n", " ") for cell in r.cells]
                rows_data.append(row_vals)

            if rows_data:
                headers = rows_data[0]
                data_rows = rows_data[1:] if len(rows_data) > 1 else []
                current_section["elements"].append({
                    "type": "table",
                    "headers": headers,
                    "rows": data_rows
                })

    # Append remaining
    if current_section["elements"]:
        current_chapter["subsections"].append(current_section)
    if current_chapter["subsections"]:
        chapters.append(current_chapter)

    if not doc_title:
        doc_title = base_name.replace("_", " ").title()

    return {
        "title": doc_title,
        "author": author or "Antigravity Office Studio",
        "date": time.strftime("%B %Y"),
        "chapters": chapters,
        "asset_dir": asset_dir
    }


# ---------------------------------------------------------------------------
# Slide Spec Synthesizer (Spec-First Model)
# ---------------------------------------------------------------------------

def synthesize_slide_spec(
    doc_tree: Dict[str, Any],
    theme: str = "thesis_blue",
    max_bullets_per_slide: int = 5,
    max_table_rows_per_slide: int = 7,
) -> Dict[str, Any]:
    """
    Transforms the extracted document tree into a declarative slide specification
    matching PPTX Invariants (PPTX_INV_01 to PPTX_INV_08).
    """
    slides: List[Dict[str, Any]] = []

    # Slide 1: Premium Title Slide
    slides.append({
        "type": "title",
        "title": doc_tree["title"],
        "subtitle": "Báo Cáo Kỹ Thuật & Tài Liệu Thuyết Minh",
        "presenter": doc_tree["author"],
        "date": doc_tree["date"]
    })

    # Slide 2: Agenda / Table of Contents (Grid Cards 4x2 or 2x2)
    chapter_list = doc_tree.get("chapters", [])
    if chapter_list:
        agenda_cards = []
        for idx, ch in enumerate(chapter_list[:8]):
            agenda_cards.append({
                "title": f"Phần {idx+1}: {ch['title'][:42]}",
                "desc": f"Gồm {len(ch['subsections'])} nội dung chi tiết"
            })
        slides.append({
            "type": "grid_cards",
            "title": "NỘI DUNG THUYẾT TRÌNH (AGENDA)",
            "subtitle": "Cấu trúc tổng quan các phần báo cáo",
            "cards": agenda_cards
        })

    # Content Slides by Chapter
    for ch_idx, ch in enumerate(chapter_list):
        ch_title = ch["title"]
        ch_label = f"PHẦN {ch_idx+1}" if not re.match(r"^(Chương|Chapter|Phần)", ch_title, re.I) else ""

        # Chapter Divider Slide
        slides.append({
            "type": "chapter",
            "chapter_num": ch_label or f"CHƯƠNG {ch_idx+1}",
            "title": ch_title,
            "subtitle": f"Tổng quan phần {ch_idx+1}"
        })

        banner_text = f"{ch_label + ': ' if ch_label else ''}{ch_title}"
        if len(banner_text) > 65:
            banner_text = banner_text[:62] + "..."

        for sec in ch["subsections"]:
            sec_title = sec["title"]
            bullets: List[str] = []

            for elem in sec["elements"]:
                etype = elem["type"]

                if etype in ("bullet", "paragraph"):
                    text = elem["text"]
                    # If paragraph is long, split into key points
                    if len(text) > 160:
                        sents = _split_into_sentences(text)
                        for s in sents:
                            if len(s) > 15:
                                bullets.append(s[:180])
                    else:
                        bullets.append(text)

                    # Flush bullets if exceeding max_bullets_per_slide (PPTX_INV_04: Zero Overflow)
                    if len(bullets) >= max_bullets_per_slide:
                        slides.append({
                            "type": "content",
                            "banner": banner_text,
                            "title": sec_title,
                            "bullets": list(bullets)
                        })
                        bullets = []

                elif etype == "table":
                    # Flush pending bullets first
                    if bullets:
                        slides.append({
                            "type": "content",
                            "banner": banner_text,
                            "title": sec_title,
                            "bullets": list(bullets)
                        })
                        bullets = []

                    headers = elem.get("headers", [])
                    all_rows = elem.get("rows", [])
                    total_rows = len(all_rows)

                    if total_rows <= max_table_rows_per_slide:
                        slides.append({
                            "type": "table",
                            "banner": banner_text,
                            "title": sec_title,
                            "subtitle": "Bảng tổng hợp số liệu",
                            "headers": headers,
                            "rows": all_rows
                        })
                    else:
                        # Smart Table Pagination (PPTX_INV_06)
                        for page_start in range(0, total_rows, max_table_rows_per_slide):
                            page_rows = all_rows[page_start:page_start + max_table_rows_per_slide]
                            page_num = (page_start // max_table_rows_per_slide) + 1
                            total_pages = (total_rows + max_table_rows_per_slide - 1) // max_table_rows_per_slide
                            slides.append({
                                "type": "table",
                                "banner": banner_text,
                                "title": f"{sec_title} (Bảng trang {page_num}/{total_pages})",
                                "subtitle": f"Dữ liệu tiếp theo ({page_start+1} - {min(page_start + len(page_rows), total_rows)}/{total_rows})",
                                "headers": headers,
                                "rows": page_rows
                            })

                elif etype == "image":
                    # Flush pending bullets first
                    if bullets:
                        slides.append({
                            "type": "content",
                            "banner": banner_text,
                            "title": sec_title,
                            "bullets": list(bullets)
                        })
                        bullets = []

                    slides.append({
                        "type": "diagram",
                        "banner": banner_text,
                        "title": sec_title,
                        "image_path": elem["path"],
                        "caption": elem.get("caption") or "Sơ đồ kiến trúc / Biểu đồ minh họa"
                    })

            # Flush remaining bullets
            if bullets:
                slides.append({
                    "type": "content",
                    "banner": banner_text,
                    "title": sec_title,
                    "bullets": list(bullets)
                })
                bullets = []

    # Final Slide: Summary & Q&A Conclusion
    slides.append({
        "type": "title",
        "title": "KẾT LUẬN & HỎI ĐÁP (Q&A)",
        "subtitle": "Xin trân trọng cảm ơn sự theo dõi của Quý vị!",
        "presenter": doc_tree["author"],
        "date": doc_tree["date"]
    })

    return {
        "spec_version": "3.0",
        "theme": theme,
        "source": doc_tree["title"],
        "total_slides": len(slides),
        "slides": slides
    }


# ---------------------------------------------------------------------------
# High-Level Conversion Pipeline
# ---------------------------------------------------------------------------

def convert_docx_to_pptx(
    docx_path: str,
    output_pptx: Optional[str] = None,
    spec_out: Optional[str] = None,
    theme: str = "thesis_blue",
    template_path: Optional[str] = None,
) -> Tuple[str, str]:
    """
    Converts a Word (.docx) document to a 16:9 PowerPoint (.pptx) presentation.

    Args:
        docx_path:      Path to source Word .docx file.
        output_pptx:    Path for generated PowerPoint .pptx file.
        spec_out:       Optional path to write the intermediate slide_spec.json.
        theme:          Theme name ('thesis_blue', 'corporate_blue', 'modern_dark', 'academic_light').
        template_path:  Optional base .pptx / .potx template for master slide inheritance.

    Returns:
        Tuple of (path_to_pptx, path_to_spec_json).
    """
    docx_path = os.path.abspath(docx_path)
    if not os.path.isfile(docx_path):
        raise FileNotFoundError(f"Source DOCX file not found: {docx_path}")

    if docx_path.lower().endswith(".pdf"):
        return convert_pdf_to_pptx(
            pdf_path=docx_path,
            output_pptx=output_pptx,
            spec_out=spec_out,
            theme=theme,
            template_path=template_path,
        )

    base_dir = os.path.dirname(docx_path)
    base_name = os.path.splitext(os.path.basename(docx_path))[0]

    if not output_pptx:
        output_pptx = os.path.join(base_dir, f"{base_name}_presentation.pptx")
    output_pptx = os.path.abspath(output_pptx)

    if not spec_out:
        spec_out = os.path.join(base_dir, f"{base_name}_slide_spec.json")
    spec_out = os.path.abspath(spec_out)

    t0 = time.time()

    # 1. Parse Word AST
    doc_tree = extract_docx_structure(docx_path)

    # 2. Synthesize Slide Spec
    spec = synthesize_slide_spec(doc_tree, theme=theme)

    # 3. Export Spec JSON
    os.makedirs(os.path.dirname(spec_out), exist_ok=True)
    with open(spec_out, "w", encoding="utf-8") as f:
        json.dump(spec, f, ensure_ascii=False, indent=2)

    # 4. Render PowerPoint
    final_pptx = write_pptx_from_spec(
        spec=spec,
        output_path=output_pptx,
        theme_name=theme,
        template_path=template_path
    )

    elapsed = time.time() - t0
    print(f"[OK] DOCX -> PPTX conversion completed in {elapsed:.2f}s")
    print(f"     Slides Generated: {len(spec['slides'])}")
    print(f"     Spec JSON:        {spec_out}")
    print(f"     PowerPoint Deck:  {final_pptx}")

    return final_pptx, spec_out


def convert_pdf_to_pptx(
    pdf_path: str,
    output_pptx: Optional[str] = None,
    spec_out: Optional[str] = None,
    theme: str = "thesis_blue",
    template_path: Optional[str] = None,
    page_range: Optional[str] = None,
) -> Tuple[str, str]:
    """
    Converts a PDF document to a 16:9 PowerPoint presentation.
    Extracts high-fidelity structure, headings, tables, and images.

    Args:
        pdf_path:       Path to source PDF file.
        output_pptx:    Path for generated PowerPoint .pptx file.
        spec_out:       Optional path to write intermediate slide_spec.json.
        theme:          Visual theme ('thesis_blue', 'corporate_blue', 'modern_dark', 'academic_light').
        template_path:  Optional base .pptx / .potx template.
        page_range:     Optional page range filter (e.g. '1-10').

    Returns:
        Tuple of (path_to_pptx, path_to_spec_json).
    """
    pdf_path = os.path.abspath(pdf_path)
    if not os.path.isfile(pdf_path):
        raise FileNotFoundError(f"PDF file not found: {pdf_path}")

    base_dir = os.path.dirname(pdf_path)
    base_name = os.path.splitext(os.path.basename(pdf_path))[0]
    temp_docx = os.path.join(base_dir, f".{base_name}_intermediate.docx")

    try:
        from .converter_engine import convert_file
    except (ImportError, ValueError):
        from converter_engine import convert_file  # type: ignore

    t0 = time.time()
    ok, actual_docx, _ = convert_file(pdf_path, temp_docx, page_range=page_range)
    if not ok:
        raise RuntimeError(f"Failed to extract structure from PDF: {actual_docx}")

    try:
        final_pptx, final_spec = convert_docx_to_pptx(
            docx_path=actual_docx,
            output_pptx=output_pptx or os.path.join(base_dir, f"{base_name}_presentation.pptx"),
            spec_out=spec_out or os.path.join(base_dir, f"{base_name}_slide_spec.json"),
            theme=theme,
            template_path=template_path,
        )
        elapsed = time.time() - t0
        print(f"[OK] PDF -> PPTX full-cycle conversion completed in {elapsed:.2f}s")
        return final_pptx, final_spec
    finally:
        if os.path.isfile(actual_docx) and actual_docx.endswith("_intermediate.docx"):
            try:
                os.remove(actual_docx)
            except Exception:
                pass
