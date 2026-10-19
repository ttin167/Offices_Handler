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
from typing import Any, Callable, Dict, List, Optional, Tuple

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

BOX_CHARS = set("┌┐└┘├┤┬┴┼─│═║╔╗╚╝╠╣╦╩╬▲▼◄►")


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

    # Pattern for Major Numbered Chapter like "1. TỔNG QUAN" or "Chương 1" or "Phần 1"
    is_numbered_chapter = bool(re.match(r"^(?:(?:[0-9]+|[IVXLCDM]+)\.\s+[A-ZÀ-Ỹ]|(?:Chương|Phần|Chapter)\s+\d+)", text, re.IGNORECASE))

    if "heading 1" in style_name or "tiêu đề 1" in style_name:
        return True, 1
    if "heading 2" in style_name or "tiêu đề 2" in style_name:
        if is_numbered_chapter:
            return True, 1
        return True, 2
    if "heading 3" in style_name or "tiêu đề 3" in style_name:
        return True, 3
    if "title" in style_name:
        return True, 1

    # Fallback heuristic: check numbering pattern
    if is_numbered_chapter and len(text) < 120:
        return True, 1
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
    raw = re.split(r"(?<=[.!?])\s+", text.strip())
    sentences = [s.strip() for s in raw if len(s.strip()) > 10]
    return sentences if sentences else [text.strip()]


def _flush_bullets(bullets_list: List[str], max_per_slide: int = 5) -> List[List[str]]:
    """Splits a list of bullets into chunks ensuring no lone 1-bullet orphan slides."""
    if not bullets_list:
        return []
    if len(bullets_list) <= max_per_slide + 1:  # e.g. 6 bullets -> 1 slide
        return [bullets_list]
    chunks: List[List[str]] = []
    for i in range(0, len(bullets_list), max_per_slide):
        chunk = bullets_list[i : i + max_per_slide]
        if len(chunk) == 1 and chunks:
            chunk = [chunks[-1].pop()] + chunk
        chunks.append(chunk)
    return chunks


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
    if author.lower() in ("python-docx", "none", "null", ""):
        author = "Antigravity Office Studio"

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
                # ASCII diagram detection
                if any(c in text for c in BOX_CHARS):
                    if current_section["elements"] and current_section["elements"][-1]["type"] == "ascii_diagram":
                        current_section["elements"][-1]["code"] += "\n" + text
                    else:
                        current_section["elements"].append({
                            "type": "ascii_diagram",
                            "code": text
                        })
                    continue

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
# Slide Spec Synthesizer & Semantic Classifier (v4.0)
# ---------------------------------------------------------------------------

def _detect_business_layout(sec_title: str, elements: List[Dict[str, Any]]) -> Optional[Dict[str, Any]]:
    """
    Analyzes section title and content elements to determine if a specialized
    business layout (SWOT, Timeline, KPI Summary, Process) applies (v4.0).
    """
    title_lower = sec_title.lower()

    # 1. SWOT Matrix Heuristic
    if "swot" in title_lower or ("điểm mạnh" in title_lower and "điểm yếu" in title_lower):
        quadrants = {"strengths": [], "weaknesses": [], "opportunities": [], "threats": []}
        current_q = "strengths"
        for el in elements:
            txt = el.get("text", "").strip()
            tl = txt.lower()
            if "điểm mạnh" in tl or "strength" in tl:
                current_q = "strengths"
            elif "điểm yếu" in tl or "weakness" in tl:
                current_q = "weaknesses"
            elif "cơ hội" in tl or "opportunit" in tl:
                current_q = "opportunities"
            elif "thách thức" in tl or "threat" in tl:
                current_q = "threats"

            if ":" in txt:
                _, val = txt.split(":", 1)
                clean_t = val.strip()
            else:
                clean_t = re.sub(r"^[-•*]\s*", "", txt).strip()

            if clean_t and not any(clean_t.lower().startswith(k) for k in ("điểm mạnh", "điểm yếu", "cơ hội", "thách thức", "strengths", "weaknesses", "opportunities", "threats")):
                quadrants[current_q].append(clean_t)

        if any(quadrants.values()):
            return {"type": "swot", "swot": quadrants}

    # 2. Timeline / Roadmap Heuristic
    if any(k in title_lower for k in ("lộ trình", "roadmap", "timeline", "kế hoạch triển khai", "milestone")):
        milestones = []
        for el in elements:
            txt = el.get("text", "")
            m_match = re.match(r"^(Q\d[\/\-]?\d*|Giai đoạn \d+|Phase \d+|Tháng \d+|Năm \d{4}|\d{4})\s*[:\-–]\s*(.*)$", txt, re.IGNORECASE)
            if m_match:
                phase_part = m_match.group(1).strip()
                desc_part = m_match.group(2).strip()
                milestones.append({
                    "phase": phase_part,
                    "title": desc_part[:40] if len(desc_part) > 40 else desc_part,
                    "tasks": [desc_part] if desc_part else ["Triển khai mục tiêu"]
                })
            else:
                clean_t = re.sub(r"^[-•*]\s*", "", txt).strip()
                if clean_t and milestones:
                    milestones[-1]["tasks"].append(clean_t)
        if len(milestones) >= 2:
            return {"type": "timeline", "milestones": milestones}

    # 3. KPI Summary Heuristic
    if any(k in title_lower for k in ("kpi", "chỉ số", "metrics", "kết quả chính", "hiệu suất", "doanh thu")):
        metrics = []
        for el in elements:
            txt = el.get("text", "")
            kpi_match = re.search(r"([^:\n]+):\s*([0-9\.,]+\s*[%kMBkK\wđ\$₫]*)(?:\s*\(([^)]+)\))?", txt)
            if kpi_match:
                lbl = kpi_match.group(1).strip().lstrip("-•* ")
                val = kpi_match.group(2).strip()
                change = kpi_match.group(3) or ""
                if any(c.isdigit() for c in val):
                    metrics.append({
                        "label": lbl,
                        "value": val,
                        "change": change,
                        "desc": txt
                    })
        if len(metrics) >= 2:
            return {"type": "kpi_summary", "metrics": metrics[:4]}

    # 4. Sequential Process / Steps Heuristic
    if any(k in title_lower for k in ("quy trình", "tiến trình", "các bước", "workflow", "process", "hướng dẫn")):
        steps = []
        for el in elements:
            txt = el.get("text", "")
            s_match = re.match(r"^(Bước \d+|Step \d+|Phase \d+|\d+\.)\s*[:\-–]?\s*(.*)$", txt, re.IGNORECASE)
            if s_match:
                step_title = s_match.group(2).strip()
                steps.append({
                    "title": step_title[:45] if len(step_title) > 45 else (step_title or f"Bước {len(steps)+1}"),
                    "actions": []
                })
            else:
                clean_t = re.sub(r"^[-•*]\s*", "", txt).strip()
                if clean_t and steps:
                    steps[-1]["actions"].append(clean_t)
        if len(steps) >= 3:
            return {"type": "process", "steps": steps[:6]}

    return None


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
    chapter_list = [ch for ch in doc_tree.get("chapters", []) if not re.search(r"\bMỤC\s+LỤC\b", ch["title"], re.IGNORECASE)]
    if chapter_list:
        agenda_cards = []
        for idx, ch in enumerate(chapter_list):
            clean_title = re.sub(r"^(?:(?:[0-9]+|[IVXLCDM]+)\.\s*|(?:Chương|Phần|Chapter)\s+\d+[:\.]?\s*)", "", ch["title"]).strip()
            agenda_cards.append({
                "title": f"Phần {idx+1}: {clean_title[:45]}",
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
        ch_title = ch["title"].strip()
        if re.search(r"\bMỤC\s+LỤC\b", ch_title, re.IGNORECASE):
            continue
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
            sec_title = sec["title"].strip()
            if re.search(r"\bMỤC\s+LỤC\b", sec_title, re.IGNORECASE):
                continue

            # Deduplicate banner when sec_title matches or heavily overlaps ch_title
            sec_banner = banner_text
            t_norm = re.sub(r'^\d+[\.\s\-]*', '', sec_title).strip().lower()
            c_norm = re.sub(r'^\d+[\.\s\-]*', '', ch_title).strip().lower()
            if t_norm and c_norm and (t_norm == c_norm or t_norm in c_norm or c_norm in t_norm):
                sec_banner = ch_label or f"PHẦN {ch_idx+1}"

            # Smart Semantic Classification for Business Layouts (v4.0)
            biz_layout = _detect_business_layout(sec_title, sec.get("elements", []))
            if biz_layout:
                biz_layout["banner"] = sec_banner
                biz_layout["title"] = sec_title
                slides.append(biz_layout)
                continue

            bullets: List[str] = []

            for elem in sec["elements"]:
                etype = elem["type"]

                if etype in ("bullet", "paragraph"):
                    text = elem["text"].strip()
                    if not text:
                        continue
                    # If paragraph is long, split into key points without hard character truncation
                    if len(text) > 160:
                        sents = _split_into_sentences(text)
                        for s in sents:
                            if len(s) > 15:
                                bullets.append(s.strip())
                    else:
                        bullets.append(text)

                    # Flush bullets if exceeding threshold
                    if len(bullets) >= max_bullets_per_slide + 1:
                        b_chunks = _flush_bullets(bullets, max_bullets_per_slide)
                        if len(b_chunks) > 1:
                            for chunk in b_chunks[:-1]:
                                slides.append({
                                    "type": "content",
                                    "banner": sec_banner,
                                    "title": sec_title,
                                    "bullets": chunk
                                })
                            bullets = list(b_chunks[-1])

                elif etype == "ascii_diagram":
                    if bullets:
                        for chunk in _flush_bullets(bullets, max_bullets_per_slide):
                            slides.append({
                                "type": "content",
                                "banner": sec_banner,
                                "title": sec_title,
                                "bullets": chunk
                            })
                        bullets = []

                    slides.append({
                        "type": "code",
                        "banner": sec_banner,
                        "title": sec_title,
                        "subtitle": "Sơ đồ kiến trúc phân tầng hệ thống",
                        "code": elem["code"],
                        "language": "text"
                    })

                elif etype == "table":
                    table_sub = "Bảng tổng hợp số liệu"
                    # If there was exactly 1 short bullet before table that ends with ':' or is brief,
                    # fold it into table subtitle to prevent an orphan slide
                    if len(bullets) == 1 and (bullets[0].endswith(":") or len(bullets[0]) < 90):
                        table_sub = bullets.pop(0).rstrip(":")
                    elif bullets:
                        for chunk in _flush_bullets(bullets, max_bullets_per_slide):
                            slides.append({
                                "type": "content",
                                "banner": sec_banner,
                                "title": sec_title,
                                "bullets": chunk
                            })
                        bullets = []

                    headers = elem.get("headers", [])
                    all_rows = elem.get("rows", [])
                    total_rows = len(all_rows)

                    if total_rows <= max_table_rows_per_slide:
                        slides.append({
                            "type": "table",
                            "banner": sec_banner,
                            "title": sec_title,
                            "subtitle": table_sub,
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
                                "banner": sec_banner,
                                "title": f"{sec_title} (Bảng trang {page_num}/{total_pages})",
                                "subtitle": f"{table_sub} ({page_start+1} - {min(page_start + len(page_rows), total_rows)}/{total_rows})",
                                "headers": headers,
                                "rows": page_rows
                            })

                elif etype == "image":
                    if bullets:
                        for chunk in _flush_bullets(bullets, max_bullets_per_slide):
                            slides.append({
                                "type": "content",
                                "banner": sec_banner,
                                "title": sec_title,
                                "bullets": chunk
                            })
                        bullets = []

                    slides.append({
                        "type": "diagram",
                        "banner": sec_banner,
                        "title": sec_title,
                        "image_path": elem["path"],
                        "caption": elem.get("caption") or "Sơ đồ kiến trúc / Biểu đồ minh họa"
                    })

            # Flush remaining bullets
            if bullets:
                for chunk in _flush_bullets(bullets, max_bullets_per_slide):
                    slides.append({
                        "type": "content",
                        "banner": sec_banner,
                        "title": sec_title,
                        "bullets": chunk
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
    theme_name: Optional[str] = None,
    progress_callback: Optional[Callable[[float, str], None]] = None,
) -> Tuple[str, str]:
    """
    Converts a Word (.docx) document to a 16:9 PowerPoint (.pptx) presentation.

    Args:
        docx_path:          Path to source Word .docx file.
        output_pptx:        Path for generated PowerPoint .pptx file.
        spec_out:           Optional path to write the intermediate slide_spec.json.
        theme:              Theme name ('thesis_blue', 'corporate_blue', 'modern_dark', 'academic_light').
        template_path:      Optional base .pptx / .potx template for master slide inheritance.
        theme_name:         Optional alias for theme.
        progress_callback:  Optional progress callback (pct, msg).

    Returns:
        Tuple of (path_to_pptx, path_to_spec_json).
    """
    if theme_name:
        theme = theme_name

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
            progress_callback=progress_callback,
        )

    if progress_callback:
        progress_callback(10.0, "Đang phân tích cấu trúc tài liệu Word...")

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

    if progress_callback:
        progress_callback(40.0, "Đang tổng hợp các slide thuyết trình...")

    # 2. Synthesize Slide Spec
    spec = synthesize_slide_spec(doc_tree, theme=theme)

    # 3. Export Spec JSON
    os.makedirs(os.path.dirname(spec_out), exist_ok=True)
    with open(spec_out, "w", encoding="utf-8") as f:
        json.dump(spec, f, ensure_ascii=False, indent=2)

    if progress_callback:
        progress_callback(70.0, "Đang kết xuất tệp PowerPoint...")

    # 4. Render PowerPoint
    final_pptx = write_pptx_from_spec(
        spec=spec,
        output_path=output_pptx,
        theme_name=theme,
        template_path=template_path
    )

    if progress_callback:
        progress_callback(100.0, "Hoàn tất chuyển đổi sang PowerPoint!")

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
    theme_name: Optional[str] = None,
    progress_callback: Optional[Callable[[float, str], None]] = None,
) -> Tuple[str, str]:
    """
    Converts a PDF document to a 16:9 PowerPoint presentation.
    Extracts high-fidelity structure, headings, tables, and images.

    Args:
        pdf_path:           Path to source PDF file.
        output_pptx:        Path for generated PowerPoint .pptx file.
        spec_out:           Optional path to write intermediate slide_spec.json.
        theme:              Visual theme ('thesis_blue', 'corporate_blue', 'modern_dark', 'academic_light').
        template_path:      Optional base .pptx / .potx template.
        page_range:         Optional page range filter (e.g. '1-10').
        theme_name:         Optional alias for theme.
        progress_callback:  Optional progress callback.

    Returns:
        Tuple of (path_to_pptx, path_to_spec_json).
    """
    if theme_name:
        theme = theme_name

    pdf_path = os.path.abspath(pdf_path)
    if not os.path.isfile(pdf_path):
        raise FileNotFoundError(f"PDF file not found: {pdf_path}")

    if progress_callback:
        progress_callback(10.0, "Đang bóc tách trang PDF...")

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
            progress_callback=progress_callback,
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
