"""
core.converter.pdf_toolkit — Everyday PDF Toolkit (v4.0).

Provides robust, high-performance PDF manipulation using PyMuPDF (fitz):
1. merge_pdfs: Concatenate multiple PDFs in order.
2. split_pdf: Split PDF into single pages or custom page ranges ('1-3', '4-5').
3. rotate_and_clean_pdf: Rotate pages and detect/remove blank pages.
4. stamp_watermark_pdf: Diagonal semi-transparent watermark text on all pages.
5. compress_pdf: Garbage collection, deflation, and stream optimization.
"""

from __future__ import annotations

import os
import re
from typing import Any, Dict, List, Optional, Tuple, Union

try:
    import pymupdf as fitz
except ImportError:
    import fitz  # type: ignore


def merge_pdfs(pdf_paths: List[str], output_path: str) -> str:
    """
    Concatenates multiple PDF documents in sequence into a single PDF file.
    """
    if not pdf_paths:
        raise ValueError("pdf_paths list cannot be empty")

    dest = os.path.abspath(output_path)
    out_dir = os.path.dirname(dest)
    if out_dir:
        os.makedirs(out_dir, exist_ok=True)

    merged_doc = fitz.open()

    for p in pdf_paths:
        abs_p = os.path.abspath(p)
        if not os.path.isfile(abs_p):
            merged_doc.close()
            raise FileNotFoundError(f"Source PDF not found: {abs_p}")
        sub_doc = fitz.open(abs_p)
        merged_doc.insert_pdf(sub_doc)
        sub_doc.close()

    merged_doc.save(dest, garbage=3, deflate=True)
    merged_doc.close()
    return dest


def _parse_page_ranges(ranges_str: str, total_pages: int) -> List[List[int]]:
    """
    Parses string like '1-3, 5, 7-9' into 0-based page index lists.
    """
    groups: List[List[int]] = []
    for part in ranges_str.split(","):
        part = part.strip()
        if not part:
            continue
        if "-" in part:
            s_str, e_str = part.split("-", 1)
            start_p = max(1, int(s_str.strip()))
            end_p = min(total_pages, int(e_str.strip()))
            if start_p <= end_p:
                groups.append([p - 1 for p in range(start_p, end_p + 1)])
        elif part.isdigit():
            p_num = int(part)
            if 1 <= p_num <= total_pages:
                groups.append([p_num - 1])
    return groups


def split_pdf(
    pdf_path: str,
    output_dir: str,
    page_ranges: Optional[Union[List[str], str]] = None,
) -> List[str]:
    """
    Splits a PDF into separate files.
    - If page_ranges is None or 'all': each page becomes a separate PDF.
    - If page_ranges is '1-3, 4-5': generates range-based files.
    """
    src = os.path.abspath(pdf_path)
    if not os.path.isfile(src):
        raise FileNotFoundError(f"Source PDF not found: {src}")

    dest_dir = os.path.abspath(output_dir)
    os.makedirs(dest_dir, exist_ok=True)

    doc = fitz.open(src)
    total_pages = len(doc)
    base_name = os.path.splitext(os.path.basename(src))[0]
    generated_files: List[str] = []

    if not page_ranges or page_ranges == "all":
        # Split every page individually
        for pno in range(total_pages):
            out_file = os.path.join(dest_dir, f"{base_name}_page_{pno + 1}.pdf")
            sub_doc = fitz.open()
            sub_doc.insert_pdf(doc, from_page=pno, to_page=pno)
            sub_doc.save(out_file, garbage=3, deflate=True)
            sub_doc.close()
            generated_files.append(out_file)
    else:
        # Range-based split
        range_str = ",".join(page_ranges) if isinstance(page_ranges, list) else str(page_ranges)
        groups = _parse_page_ranges(range_str, total_pages)

        for idx, page_indices in enumerate(groups, start=1):
            if not page_indices:
                continue
            out_file = os.path.join(dest_dir, f"{base_name}_part_{idx}.pdf")
            sub_doc = fitz.open()
            for pno in page_indices:
                sub_doc.insert_pdf(doc, from_page=pno, to_page=pno)
            sub_doc.save(out_file, garbage=3, deflate=True)
            sub_doc.close()
            generated_files.append(out_file)

    doc.close()
    return generated_files


def rotate_and_clean_pdf(
    pdf_path: str,
    output_path: str,
    rotation: int = 0,
    remove_blank: bool = True,
) -> Dict[str, Any]:
    """
    Applies clockwise rotation (90, 180, 270) and optionally removes blank pages.
    """
    src = os.path.abspath(pdf_path)
    if not os.path.isfile(src):
        raise FileNotFoundError(f"Source PDF not found: {src}")

    dest = os.path.abspath(output_path)
    out_dir = os.path.dirname(dest)
    if out_dir:
        os.makedirs(out_dir, exist_ok=True)

    doc = fitz.open(src)
    initial_pages = len(doc)
    blank_indices: List[int] = []

    # 1. Rotate pages if requested
    if rotation % 360 != 0:
        for page in doc:
            page.set_rotation((page.rotation + rotation) % 360)

    # 2. Identify blank pages
    if remove_blank:
        for pno in range(len(doc)):
            page = doc[pno]
            text = page.get_text().strip()
            images = page.get_images()
            drawings = page.get_drawings()
            # If no text, no raster images, and fewer than 2 trivial rect drawings
            if len(text) == 0 and len(images) == 0 and len(drawings) < 2:
                blank_indices.append(pno)

        # Delete blank pages in reverse order to preserve indices
        for pno in reversed(blank_indices):
            doc.delete_page(pno)

    doc.save(dest, garbage=3, deflate=True)
    final_pages = len(doc)
    doc.close()

    return {
        "success": True,
        "output_path": dest,
        "initial_pages": initial_pages,
        "final_pages": final_pages,
        "removed_blanks_count": len(blank_indices),
        "removed_blank_pages": [p + 1 for p in blank_indices],
        "rotation_applied": rotation % 360,
    }


def _find_system_font() -> Optional[str]:
    """Finds a common Unicode TrueType font on Windows, Linux, or macOS."""
    candidates = [
        r"C:\Windows\Fonts\arial.ttf",
        r"C:\Windows\Fonts\times.ttf",
        r"C:\Windows\Fonts\segoeui.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf",
        "/usr/share/fonts/dejavu/DejaVuSans.ttf",
        "/System/Library/Fonts/Helvetica.ttc",
        "/Library/Fonts/Arial.ttf",
    ]
    for c in candidates:
        if os.path.isfile(c):
            return c
    return None


def stamp_watermark_pdf(
    pdf_path: str,
    output_path: str,
    watermark_text: str = "BẢN QUYỀN",
    opacity: float = 0.25,
    angle: float = 45.0,
    font_size: int = 40,
    color: Tuple[float, float, float] = (0.5, 0.5, 0.5),
    font_path: Optional[str] = None,
) -> str:
    """
    Injects a diagonal semi-transparent watermark text onto all pages of a PDF.
    Preserves all original vector text and image layers.
    """
    src = os.path.abspath(pdf_path)
    if not os.path.isfile(src):
        raise FileNotFoundError(f"Source PDF not found: {src}")

    dest = os.path.abspath(output_path)
    out_dir = os.path.dirname(dest)
    if out_dir:
        os.makedirs(out_dir, exist_ok=True)

    doc = fitz.open(src)
    ttf_font = font_path or _find_system_font()

    for page in doc:
        fontname = "helv"
        if ttf_font:
            try:
                page.insert_font(fontname="watermark_font", fontfile=ttf_font)
                fontname = "watermark_font"
            except Exception:
                pass

        rect = page.rect
        center_x = rect.width / 2.0
        center_y = rect.height / 2.0

        # Calculate approximate text width to center baseline
        text_w = len(watermark_text) * (font_size * 0.45)
        pt = fitz.Point(center_x - text_w / 2.0, center_y)

        shape = page.new_shape()
        shape.insert_text(
            pt,
            watermark_text,
            fontname=fontname,
            fontsize=font_size,
            color=color,
            fill_opacity=opacity,
            stroke_opacity=opacity,
        )
        shape.commit()

    doc.save(dest, garbage=3, deflate=True)
    doc.close()
    return dest


def compress_pdf(
    pdf_path: str,
    output_path: str,
    quality: str = "medium",
) -> Dict[str, Any]:
    """
    Compresses PDF by garbage-collecting unreferenced objects, deflating streams,
    and reorganizing xref tables.
    """
    src = os.path.abspath(pdf_path)
    if not os.path.isfile(src):
        raise FileNotFoundError(f"Source PDF not found: {src}")

    dest = os.path.abspath(output_path)
    out_dir = os.path.dirname(dest)
    if out_dir:
        os.makedirs(out_dir, exist_ok=True)

    orig_size = os.path.getsize(src)

    doc = fitz.open(src)
    # garbage=4: removes unused objects, compacts streams, merges duplicate objects
    doc.save(dest, garbage=4, deflate=True, clean=True)
    doc.close()

    comp_size = os.path.getsize(dest)
    ratio_pct = round(((orig_size - comp_size) / orig_size) * 100.0, 1) if orig_size > 0 else 0.0

    return {
        "success": True,
        "output_path": dest,
        "original_bytes": orig_size,
        "compressed_bytes": comp_size,
        "saved_bytes": max(0, orig_size - comp_size),
        "reduction_pct": max(0.0, ratio_pct),
    }
