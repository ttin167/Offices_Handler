"""
core.converter.pdf_toolkit — Everyday PDF Toolkit (v4.0).

Provides robust, high-performance PDF manipulation using PyMuPDF (fitz):
1. merge_pdfs: Concatenate multiple PDFs in order.
2. split_pdf: Split PDF into single pages or custom page ranges ('1-3', '4-5').
3. rotate_and_clean_pdf: Rotate pages and detect/remove blank pages.
4. stamp_watermark_pdf: Diagonal semi-transparent watermark text on all pages.
5. compress_pdf: Garbage collection, deflation, and stream optimization.
"""

import base64
import io
import math
import os
import re
import zipfile
from typing import Any, Dict, List, Optional, Tuple, Union

from PIL import Image, ImageOps

try:
    import openpyxl
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter
except ImportError:
    openpyxl = None  # type: ignore

try:
    import pymupdf as fitz
except ImportError:
    import fitz  # type: ignore

try:
    from core.xlsx.writer import apply_worksheet_print_setup
except ImportError:
    apply_worksheet_print_setup = None  # type: ignore


# Built-in Regex Lexicon for Vietnamese & International Sensitive Information
SENSITIVE_PATTERNS: Dict[str, str] = {
    "cccd": r"\b0\d{11}\b|\b\d{9}\b",
    "phone": r"(?:\+84|0)(?:3[2-9]|5[2689]|7[06-9]|8[1-9]|9[0-9])\d{7}\b",
    "email": r"\b[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+\b",
    "bank_card": r"\b(?:\d{4}[ -]?){3}\d{4}\b|\b\d{10,19}\b",
    "salary": r"\b\d{1,3}(?:\.\d{3}){1,3}\s*(?:VNĐ|VND|đ| đồng)\b",
}


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


def redact_sensitive_pdf(
    pdf_path: str,
    output_path: str,
    patterns: Optional[List[str]] = None,
    custom_terms: Optional[List[str]] = None,
    rect_list: Optional[List[Dict[str, Any]]] = None,
    fill_color: Tuple[float, float, float] = (0.0, 0.0, 0.0),
) -> Dict[str, Any]:
    """
    Permanently redacts sensitive information (CCCD, phone, email, bank card,
    custom keywords, or specified bounding boxes) from a PDF document.
    Deletes underlying character bytes and raster image pixels (True Redaction).
    """
    src = os.path.abspath(pdf_path)
    if not os.path.isfile(src):
        raise FileNotFoundError(f"Source PDF not found: {src}")

    dest = os.path.abspath(output_path)
    out_dir = os.path.dirname(dest)
    if out_dir:
        os.makedirs(out_dir, exist_ok=True)

    doc = fitz.open(src)
    total_redactions = 0
    pages_affected: set[int] = set()
    matched_categories: Dict[str, int] = {}

    # Compile regex patterns
    active_patterns: List[Tuple[str, re.Pattern]] = []
    if patterns:
        for p in patterns:
            key = p.lower().strip()
            regex_str = SENSITIVE_PATTERNS.get(key, p)
            try:
                active_patterns.append((key, re.compile(regex_str, re.IGNORECASE)))
            except re.error:
                continue

    # Process each page
    for pno in range(len(doc)):
        page = doc[pno]
        page_text = page.get_text()
        page_had_redaction = False

        # 1. Search by Regex patterns
        for cat_name, rx in active_patterns:
            matches = rx.findall(page_text)
            for m in matches:
                m_str = m.strip() if isinstance(m, str) else str(m)
                if not m_str or len(m_str) < 3:
                    continue
                rects = page.search_for(m_str)
                for r in rects:
                    page.add_redact_annot(r, fill=fill_color)
                    total_redactions += 1
                    page_had_redaction = True
                    matched_categories[cat_name] = matched_categories.get(cat_name, 0) + 1

        # 2. Search by custom terms
        if custom_terms:
            for term in custom_terms:
                t_clean = term.strip()
                if not t_clean:
                    continue
                rects = page.search_for(t_clean)
                for r in rects:
                    page.add_redact_annot(r, fill=fill_color)
                    total_redactions += 1
                    page_had_redaction = True
                    matched_categories["custom_terms"] = matched_categories.get("custom_terms", 0) + 1

        # 3. Explicit coordinate rects
        if rect_list:
            for item in rect_list:
                item_page = item.get("page", 1) - 1
                if item_page == pno:
                    r = fitz.Rect(item["x0"], item["y0"], item["x1"], item["y1"])
                    page.add_redact_annot(r, fill=fill_color)
                    total_redactions += 1
                    page_had_redaction = True
                    matched_categories["coordinates"] = matched_categories.get("coordinates", 0) + 1

        # Apply redactions to cut both text and underlying image pixels
        if page_had_redaction:
            pages_affected.add(pno + 1)
            try:
                page.apply_redactions(images=fitz.PDF_REDACT_IMAGE_PIXELS)
            except (AttributeError, TypeError):
                page.apply_redactions()

    doc.save(dest, garbage=4, deflate=True, clean=True)
    doc.close()

    return {
        "success": True,
        "output_path": dest,
        "total_redactions": total_redactions,
        "pages_affected_count": len(pages_affected),
        "pages_affected": sorted(list(pages_affected)),
        "categories_breakdown": matched_categories,
    }


def protect_pdf(
    pdf_path: str,
    output_path: str,
    user_password: Optional[str] = None,
    owner_password: Optional[str] = None,
    allow_print: bool = True,
    allow_copy: bool = False,
    allow_modify: bool = False,
    allow_annotate: bool = False,
) -> Dict[str, Any]:
    """
    Encrypts a PDF using military-grade AES-256 encryption with separate User/Owner
    passwords and explicit feature permissions.
    """
    src = os.path.abspath(pdf_path)
    if not os.path.isfile(src):
        raise FileNotFoundError(f"Source PDF not found: {src}")

    dest = os.path.abspath(output_path)
    out_dir = os.path.dirname(dest)
    if out_dir:
        os.makedirs(out_dir, exist_ok=True)

    doc = fitz.open(src)

    # Calculate standard permission bits
    perm = 0
    if allow_print:
        perm |= fitz.PDF_PERM_PRINT
    if allow_copy:
        perm |= fitz.PDF_PERM_COPY
    if allow_modify:
        perm |= fitz.PDF_PERM_MODIFY
    if allow_annotate:
        perm |= fitz.PDF_PERM_ANNOTATE

    u_pw = user_password or ""
    o_pw = owner_password or user_password or "AdminSecret#2026"

    doc.save(
        dest,
        encryption=fitz.PDF_ENCRYPT_AES_256,
        user_pw=u_pw,
        owner_pw=o_pw,
        permissions=perm,
        garbage=3,
        deflate=True,
    )
    doc.close()

    return {
        "success": True,
        "output_path": dest,
        "encryption": "AES-256",
        "has_user_password": bool(user_password),
        "has_owner_password": bool(owner_password),
        "permissions": {
            "allow_print": allow_print,
            "allow_copy": allow_copy,
            "allow_modify": allow_modify,
            "allow_annotate": allow_annotate,
        },
    }


def normalize_and_number_pdf(
    pdf_path: str,
    output_path: str,
    target_size: str = "A4",
    orientation: str = "auto",
    margin: float = 24.0,
    bates_prefix: Optional[str] = None,
    bates_pattern: str = "{prefix} | Trang {page:02d}/{total:02d}",
    start_number: int = 1,
    position: str = "bottom-right",
    font_size: int = 9,
    font_path: Optional[str] = None,
) -> Dict[str, Any]:
    """
    Normalizes all pages of a PDF to standard A4 (595.28 x 841.89 pt) using vector
    embedding (zero raster blur) while preserving aspect ratio, and optionally stamps
    Vietnamese Bates Numbers (e.g. 'HD-2026-001 | Trang 01/25').
    """
    src = os.path.abspath(pdf_path)
    if not os.path.isfile(src):
        raise FileNotFoundError(f"Source PDF not found: {src}")

    dest = os.path.abspath(output_path)
    out_dir = os.path.dirname(dest)
    if out_dir:
        os.makedirs(out_dir, exist_ok=True)

    src_doc = fitz.open(src)
    total_pages = len(src_doc)
    new_doc = fitz.open()

    A4_W, A4_H = 595.28, 841.89
    ttf_font = font_path or _find_system_font()

    for pno in range(total_pages):
        src_page = src_doc[pno]
        sw = src_page.rect.width
        sh = src_page.rect.height

        # Determine target A4 orientation
        if orientation == "portrait":
            tw, th = A4_W, A4_H
        elif orientation == "landscape":
            tw, th = A4_H, A4_W
        else:  # auto
            if sw > sh:
                tw, th = A4_H, A4_W
            else:
                tw, th = A4_W, A4_H

        # Calculate bounding box with margins
        avail_w = max(100.0, tw - 2 * margin)
        avail_h = max(100.0, th - 2 * margin)

        scale = min(avail_w / max(1.0, sw), avail_h / max(1.0, sh))
        dw = sw * scale
        dh = sh * scale

        x0 = margin + (avail_w - dw) / 2.0
        y0 = margin + (avail_h - dh) / 2.0
        target_rect = fitz.Rect(x0, y0, x0 + dw, y0 + dh)

        # Create new vector page and embed source page
        new_page = new_doc.new_page(width=tw, height=th)
        new_page.show_pdf_page(target_rect, src_doc, pno)

        # Apply Bates Numbering if requested
        if bates_prefix:
            curr_p = pno + start_number
            bates_text = bates_pattern.format(prefix=bates_prefix, page=curr_p, total=total_pages)

            fontname = "helv"
            if ttf_font:
                try:
                    new_page.insert_font(fontname="bates_font", fontfile=ttf_font)
                    fontname = "bates_font"
                except Exception:
                    pass

            text_w = len(bates_text) * (font_size * 0.48)

            if position == "bottom-right":
                pt = fitz.Point(tw - margin - text_w, th - (margin * 0.45))
            elif position == "bottom-center":
                pt = fitz.Point((tw - text_w) / 2.0, th - (margin * 0.45))
            elif position == "top-right":
                pt = fitz.Point(tw - margin - text_w, margin * 0.75)
            elif position == "top-left":
                pt = fitz.Point(margin, margin * 0.75)
            else:
                pt = fitz.Point(tw - margin - text_w, th - (margin * 0.45))

            shape = new_page.new_shape()
            shape.insert_text(
                pt,
                bates_text,
                fontname=fontname,
                fontsize=font_size,
                color=(0.2, 0.25, 0.3),
            )
            shape.commit()

    src_doc.close()
    new_doc.save(dest, garbage=3, deflate=True)
    new_doc.close()

    return {
        "success": True,
        "output_path": dest,
        "total_pages": total_pages,
        "target_size": target_size,
        "bates_stamped": bool(bates_prefix),
    }


def stamp_giap_lai_pdf(
    pdf_path: str,
    output_path: str,
    stamp_image_path: str,
    side: str = "right",
    stamp_diameter: float = 120.0,
    overlap_pct: float = 8.0,
) -> Dict[str, Any]:
    """
    Applies a traditional Vietnamese 'Dấu Giáp Lai' by slicing a company stamp image
    into N vertical strips and stamping them across consecutive pages at the outer margin.
    """
    src = os.path.abspath(pdf_path)
    if not os.path.isfile(src):
        raise FileNotFoundError(f"Source PDF not found: {src}")

    stamp_src = os.path.abspath(stamp_image_path)
    if not os.path.isfile(stamp_src):
        raise FileNotFoundError(f"Stamp image not found: {stamp_src}")

    dest = os.path.abspath(output_path)
    out_dir = os.path.dirname(dest)
    if out_dir:
        os.makedirs(out_dir, exist_ok=True)

    doc = fitz.open(src)
    total_pages = len(doc)
    if total_pages == 0:
        doc.close()
        raise ValueError("Cannot apply Giáp Lai stamp on an empty document")

    stamp_img = Image.open(stamp_src).convert("RGBA")
    img_w, img_h = stamp_img.size

    if total_pages == 1:
        # Single page: place stamp at edge
        buf = io.BytesIO()
        stamp_img.save(buf, format="PNG")
        page = doc[0]
        slice_w = stamp_diameter
        slice_h = stamp_diameter * (float(img_h) / float(img_w))
        y0 = (page.rect.height - slice_h) / 2.0
        x0 = page.rect.width - slice_w if side == "right" else 0.0
        rect = fitz.Rect(x0, y0, x0 + slice_w, y0 + slice_h)
        page.insert_image(rect, stream=buf.getvalue())
    else:
        # Multi-page slicing across N pages
        base_slice_w = img_w / float(total_pages)
        overlap_px = base_slice_w * (overlap_pct / 100.0)

        for i in range(total_pages):
            page = doc[i]
            x_start = max(0, int(i * base_slice_w - (overlap_px if i > 0 else 0)))
            x_end = min(img_w, int((i + 1) * base_slice_w + (overlap_px if i < total_pages - 1 else 0)))
            if x_end <= x_start:
                continue

            slice_box = (x_start, 0, x_end, img_h)
            slice_crop = stamp_img.crop(slice_box)

            buf = io.BytesIO()
            slice_crop.save(buf, format="PNG")
            slice_bytes = buf.getvalue()

            slice_disp_w = stamp_diameter * ((x_end - x_start) / float(img_w))
            slice_disp_h = stamp_diameter * (float(img_h) / float(img_w))

            page_h = page.rect.height
            page_w = page.rect.width
            y0 = (page_h - slice_disp_h) / 2.0
            y1 = y0 + slice_disp_h

            if side == "right":
                x1 = page_w
                x0 = page_w - slice_disp_w
            else:
                x0 = 0.0
                x1 = slice_disp_w

            rect = fitz.Rect(x0, y0, x1, y1)
            page.insert_image(rect, stream=slice_bytes)

    doc.save(dest, garbage=3, deflate=True)
    doc.close()

    return {
        "success": True,
        "output_path": dest,
        "total_pages": total_pages,
        "side": side,
        "stamp_slices": total_pages,
        "slices_created": total_pages,
    }


def extract_pdf_tables_to_xlsx(
    pdf_path: str,
    output_path: str,
    merge_same_schema: bool = True,
) -> Dict[str, Any]:
    """
    Extracts all structured tables from a PDF using PyMuPDF's C++ table finder (page.find_tables())
    and writes them into a beautifully styled Excel (.xlsx) file using openpyxl.
    Enforces ERR_XLSX_007 (string number format for IDs/codes) and ERR_XLSX_008 (A4 print setup).
    """
    if openpyxl is None:
        raise ImportError("openpyxl is required to extract PDF tables to Excel")

    src = os.path.abspath(pdf_path)
    if not os.path.isfile(src):
        raise FileNotFoundError(f"Source PDF not found: {src}")

    dest = os.path.abspath(output_path)
    out_dir = os.path.dirname(dest)
    if out_dir:
        os.makedirs(out_dir, exist_ok=True)

    doc = fitz.open(src)
    wb = openpyxl.Workbook()
    wb.remove(wb.active)  # Remove default blank sheet

    header_fill = PatternFill(start_color="0F294A", end_color="0F294A", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    data_font = Font(name="Calibri", size=10)
    alt_fill = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
    white_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
    thin_border = Border(
        left=Side(style="thin", color="CBD5E1"),
        right=Side(style="thin", color="CBD5E1"),
        top=Side(style="thin", color="CBD5E1"),
        bottom=Side(style="thin", color="CBD5E1"),
    )

    total_tables = 0
    total_rows = 0

    for pno in range(len(doc)):
        page = doc[pno]
        tabs = page.find_tables()
        if not tabs or len(tabs.tables) == 0:
            continue

        for t_idx, tab in enumerate(tabs.tables, start=1):
            raw_rows = tab.extract()
            if not raw_rows or len(raw_rows) == 0:
                continue

            total_tables += 1
            sheet_title = f"P{pno + 1}_Table_{t_idx}"[:31]
            ws = wb.create_sheet(title=sheet_title)

            col_widths: Dict[int, int] = {}

            for r_idx, row in enumerate(raw_rows, start=1):
                total_rows += 1
                is_header = (r_idx == 1)

                for c_idx, val in enumerate(row, start=1):
                    val_str = str(val).strip() if val is not None else ""
                    cell = ws.cell(row=r_idx, column=c_idx)

                    # Check for identifier numbers to enforce ERR_XLSX_007 (preserve leading zero)
                    is_leading_zero_code = bool(re.match(r"^0\d{3,}$", val_str))
                    is_card_or_cccd = bool(re.match(r"^\d{9,19}$", val_str))

                    if is_leading_zero_code or is_card_or_cccd:
                        cell.number_format = "@"
                        cell.value = val_str
                    elif not is_header and re.match(r"^-?\d+(?:\.\d+)?$", val_str):
                        try:
                            cell.value = float(val_str) if "." in val_str else int(val_str)
                        except ValueError:
                            cell.value = val_str
                    else:
                        cell.value = val_str

                    # Styling
                    if is_header:
                        cell.fill = header_fill
                        cell.font = header_font
                        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                    else:
                        cell.fill = alt_fill if r_idx % 2 == 0 else white_fill
                        cell.font = data_font
                        align_h = "right" if isinstance(cell.value, (int, float)) else "left"
                        cell.alignment = Alignment(horizontal=align_h, vertical="center", wrap_text=True)

                    cell.border = thin_border
                    col_widths[c_idx] = max(col_widths.get(c_idx, 10), min(50, len(val_str) + 3))

            # Apply auto column widths
            for c_idx, w in col_widths.items():
                col_letter = get_column_letter(c_idx)
                ws.column_dimensions[col_letter].width = max(12, w)

            # Apply print setup (ERR_XLSX_008)
            ws.page_setup.paperSize = ws.PAPERSIZE_A4
            if apply_worksheet_print_setup:
                apply_worksheet_print_setup(ws, orientation="landscape")
            else:
                ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE

    doc.close()

    # If no tables found, create an informative summary sheet
    if total_tables == 0:
        ws = wb.create_sheet(title="No_Tables_Detected")
        ws.cell(row=1, column=1, value="Không tìm thấy bảng biểu có cấu trúc lưới trong tệp PDF nguồn.")
        ws.cell(row=2, column=1, value="Gợi ý: Kiểm tra xem PDF có phải là bản scan thuần ảnh không (cần chạy OCR trước).")
        ws.page_setup.paperSize = ws.PAPERSIZE_A4
        if apply_worksheet_print_setup:
            apply_worksheet_print_setup(ws, orientation="landscape")
        else:
            ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE

    wb.save(dest)

    return {
        "success": True,
        "output_path": dest,
        "tables_extracted": total_tables,
        "total_rows": total_rows,
    }


def render_pdf_thumbnails(
    pdf_path: str,
    output_dir: str,
    max_pages: int = 100,
    dpi: int = 120,
) -> List[Dict[str, Any]]:
    """
    Renders high-speed thumbnail PNGs for pages of a PDF to power the Visual Page Organizer.
    """
    src = os.path.abspath(pdf_path)
    if not os.path.isfile(src):
        raise FileNotFoundError(f"Source PDF not found: {src}")

    dest_dir = os.path.abspath(output_dir)
    os.makedirs(dest_dir, exist_ok=True)

    base_name = os.path.splitext(os.path.basename(src))[0]
    doc = fitz.open(src)
    thumbnails: List[Dict[str, Any]] = []

    matrix = fitz.Matrix(dpi / 72.0, dpi / 72.0)
    pages_to_render = min(len(doc), max_pages)

    for pno in range(pages_to_render):
        page = doc[pno]
        pix = page.get_pixmap(matrix=matrix, alpha=False)
        fname = f"pdf_thumb_{base_name}_p{pno + 1}.png"
        out_file = os.path.join(dest_dir, fname)
        pix.save(out_file)

        thumbnails.append({
            "page_num": pno + 1,
            "page": pno + 1,
            "width": round(page.rect.width, 1),
            "height": round(page.rect.height, 1),
            "rotation": page.rotation,
            "filename": fname,
            "preview_url": f"/api/preview/{fname}",
            "thumbnail_url": f"/api/preview/{fname}",
            "file_path": out_file,
        })

    doc.close()
    return thumbnails


def reorganize_pdf_pages(
    pdf_path: str,
    output_path: str,
    pages_plan: List[Dict[str, Any]],
) -> Dict[str, Any]:
    """
    Executes a visual page reorganization plan:
    - pages_plan: List of dicts e.g. [{"page": 1, "rotation": 90}, {"is_blank": True}]
    - Pages not in the plan are omitted (deleted).
    - Rotations are applied per-page.
    - Order in pages_plan dictates output page sequence.
    - Supports inserting standard blank A4 pages when is_blank is True.
    """
    src = os.path.abspath(pdf_path)
    if not os.path.isfile(src):
        raise FileNotFoundError(f"Source PDF not found: {src}")

    if not pages_plan:
        raise ValueError("pages_plan cannot be empty")

    dest = os.path.abspath(output_path)
    out_dir = os.path.dirname(dest)
    if out_dir:
        os.makedirs(out_dir, exist_ok=True)

    src_doc = fitz.open(src)
    total_src_pages = len(src_doc)
    new_doc = fitz.open()

    for item in pages_plan:
        if item.get("is_blank"):
            w = float(item.get("width", 595.32))
            h = float(item.get("height", 841.92))
            new_doc.new_page(width=w, height=h)
            continue

        pno = int(item.get("page") if item.get("page") is not None else item.get("page_num", 1)) - 1
        if 0 <= pno < total_src_pages:
            new_doc.insert_pdf(src_doc, from_page=pno, to_page=pno)
            last_pno = len(new_doc) - 1
            new_page = new_doc[last_pno]
            rot = int(item.get("rotation") if item.get("rotation") is not None else item.get("rotate", 0))
            if rot % 360 != 0:
                new_page.set_rotation((new_page.rotation + rot) % 360)

    src_doc.close()
    new_doc.save(dest, garbage=3, deflate=True)
    out_pages = len(new_doc)
    new_doc.close()

    return {
        "success": True,
        "output_path": dest,
        "input_pages": total_src_pages,
        "output_pages": out_pages,
        "pages_count": out_pages,
    }


def convert_images_to_pdf(
    image_paths: List[str],
    output_path: str,
    page_size: str = "A4",
    orientation: str = "auto",
    margin_pt: float = 36.0,
) -> Dict[str, Any]:
    """
    Assembles multiple raster images (.jpg, .jpeg, .png, .webp) into a single publication-grade PDF.
    - Handles EXIF rotation automatically via ImageOps.exif_transpose.
    - page_size: "A4" (595.32 x 841.92 pt), "Letter", or "fit_image"
    - orientation: "auto" (picks landscape/portrait per image ratio), "portrait", or "landscape"
    - margin_pt: padding around the image (default 36pt = 0.5 inch).
    """
    if not image_paths:
        raise ValueError("Danh sách ảnh không được để trống")

    dest = os.path.abspath(output_path)
    out_dir = os.path.dirname(dest)
    if out_dir:
        os.makedirs(out_dir, exist_ok=True)

    doc = fitz.open()
    A4_W, A4_H = 595.32, 841.92
    processed_count = 0

    for img_p in image_paths:
        img_abs = os.path.abspath(img_p)
        if not os.path.isfile(img_abs):
            continue

        # Open with PIL to read dimensions and fix EXIF orientation
        with Image.open(img_abs) as pil_img:
            pil_img = ImageOps.exif_transpose(pil_img)
            img_w, img_h = pil_img.size

            buf = io.BytesIO()
            fmt = "PNG" if pil_img.mode in ("RGBA", "LA", "P") else "JPEG"
            if fmt == "JPEG" and pil_img.mode not in ("RGB", "L"):
                pil_img = pil_img.convert("RGB")
            pil_img.save(buf, format=fmt, quality=95)
            img_bytes = buf.getvalue()

        # Determine page dimension
        if page_size.upper() == "FIT_IMAGE":
            pw, ph = float(img_w), float(img_h)
            target_rect = fitz.Rect(0, 0, pw, ph)
        else:
            is_landscape = False
            if orientation == "landscape":
                is_landscape = True
            elif orientation == "portrait":
                is_landscape = False
            else:  # "auto"
                is_landscape = img_w > img_h

            pw = A4_H if is_landscape else A4_W
            ph = A4_W if is_landscape else A4_H

            # Calculate contained bounding box with margin
            avail_w = max(10.0, pw - 2 * margin_pt)
            avail_h = max(10.0, ph - 2 * margin_pt)

            scale = min(avail_w / img_w, avail_h / img_h)
            box_w = img_w * scale
            box_h = img_h * scale

            # Center on page
            box_x = margin_pt + (avail_w - box_w) / 2.0
            box_y = margin_pt + (avail_h - box_h) / 2.0
            target_rect = fitz.Rect(box_x, box_y, box_x + box_w, box_y + box_h)

        page = doc.new_page(width=pw, height=ph)
        page.insert_image(target_rect, stream=img_bytes, keep_proportion=True)
        processed_count += 1

    if processed_count == 0:
        doc.close()
        raise FileNotFoundError("Không có tệp ảnh hợp lệ nào được tìm thấy")

    doc.save(dest, garbage=3, deflate=True)
    total_pages = len(doc)
    doc.close()

    return {
        "success": True,
        "output_path": dest,
        "filename": os.path.basename(dest),
        "total_images": processed_count,
        "total_pages": total_pages,
    }


def export_pdf_to_images_zip(
    pdf_path: str,
    output_zip_path: str,
    dpi: int = 300,
    img_format: str = "png",
) -> Dict[str, Any]:
    """
    Renders all pages of a PDF to high-resolution images and packages them into a clean ZIP archive.
    """
    src = os.path.abspath(pdf_path)
    if not os.path.isfile(src):
        raise FileNotFoundError(f"Source PDF not found: {src}")

    dest = os.path.abspath(output_zip_path)
    out_dir = os.path.dirname(dest)
    if out_dir:
        os.makedirs(out_dir, exist_ok=True)

    doc = fitz.open(src)
    total_pages = len(doc)
    base_name = os.path.splitext(os.path.basename(src))[0]

    matrix = fitz.Matrix(dpi / 72.0, dpi / 72.0)
    fmt = img_format.lower()
    if fmt not in ("png", "jpg", "jpeg"):
        fmt = "png"

    with zipfile.ZipFile(dest, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for pno in range(total_pages):
            page = doc[pno]
            pix = page.get_pixmap(matrix=matrix, alpha=False)
            img_bytes = pix.tobytes(fmt)
            arcname = f"{base_name}_trang_{pno+1:02d}.{fmt}"
            zf.writestr(arcname, img_bytes)

    doc.close()
    return {
        "success": True,
        "output_zip": dest,
        "filename": os.path.basename(dest),
        "total_pages": total_pages,
        "dpi": dpi,
    }


def _resolve_signature_pages(page_spec: Any, total_pages: int) -> List[int]:
    """
    Resolves diverse page specifications into valid 0-based page indices.
    Accepts: 'last' (default), 'first', 'all', int, list of ints, or range strings '1, 3, 5'.
    """
    if total_pages <= 0:
        return []
    if page_spec is None or page_spec == "" or page_spec == "last" or page_spec == -1:
        return [total_pages - 1]
    if page_spec == "first":
        return [0]
    if page_spec == "all":
        return list(range(total_pages))
    if isinstance(page_spec, int):
        idx = page_spec - 1 if page_spec > 0 else total_pages + page_spec
        return [idx] if 0 <= idx < total_pages else []
    if isinstance(page_spec, (list, tuple)):
        res = []
        for p in page_spec:
            res.extend(_resolve_signature_pages(p, total_pages))
        return sorted(list(set(res)))

    s = str(page_spec).strip().lower()
    if s == "last":
        return [total_pages - 1]
    if s == "all":
        return list(range(total_pages))
    if s == "first":
        return [0]

    res = []
    for part in re.split(r"[,;]+", s):
        part = part.strip()
        if not part:
            continue
        if "-" in part:
            subparts = part.split("-", 1)
            try:
                start_p = int(subparts[0].strip())
                end_p = int(subparts[1].strip())
                for p in range(start_p, end_p + 1):
                    if 1 <= p <= total_pages:
                        res.append(p - 1)
            except ValueError:
                pass
        else:
            try:
                p = int(part)
                if 1 <= p <= total_pages:
                    res.append(p - 1)
            except ValueError:
                pass
    return sorted(list(set(res))) if res else [total_pages - 1]


def sign_pdf_with_stamp(
    pdf_path: str,
    output_path: str,
    signatures: List[Dict[str, Any]],
) -> Dict[str, Any]:
    """
    Overlays electronic signatures and company seals onto specified PDF pages at exact coordinates.
    - Handles rotation compensation (page.rotation != 0) via derotation_matrix.
    - Accepts 'image_path' or 'image_base64' (data URL or raw base64).
    - Coordinates can be passed as:
      - x_pct, y_pct, w_pct, h_pct (0.0 to 1.0 of visual page dimensions)
      - rect: [x0, y0, x1, y1] in points.
    - Flexible page selection: 'last' (default), 'first', 'all', int, or custom page list '1, 3, 5'.
    - Optional signer metadata text annotations ('signer_name', 'signer_title', 'signer_date').
    """
    src = os.path.abspath(pdf_path)
    if not os.path.isfile(src):
        raise FileNotFoundError(f"Source PDF not found: {src}")

    if not signatures:
        raise ValueError("Danh sách chữ ký / con dấu không được để trống")

    dest = os.path.abspath(output_path)
    out_dir = os.path.dirname(dest)
    if out_dir:
        os.makedirs(out_dir, exist_ok=True)

    doc = fitz.open(src)
    total_pages = len(doc)
    applied_count = 0
    ttf_font = _find_system_font()

    for sig in signatures:
        page_spec = sig.get("pages") if sig.get("pages") is not None else sig.get("page", "last")
        target_pnos = _resolve_signature_pages(page_spec, total_pages)

        # Retrieve image content once per signature spec
        img_bytes = None
        img_file = None

        if sig.get("image_base64"):
            b64_str = sig["image_base64"]
            if "," in b64_str:
                b64_str = b64_str.split(",", 1)[1]
            try:
                img_bytes = base64.b64decode(b64_str)
            except Exception as e:
                print(f"[WARN] Failed to decode base64 signature: {e}")
        elif sig.get("image_path") or sig.get("image") or sig.get("path"):
            cand = sig.get("image_path") or sig.get("image") or sig.get("path")
            cand_p = os.path.abspath(cand)
            if os.path.isfile(cand_p):
                img_file = cand_p

        if not img_bytes and not img_file:
            continue

        # Extract text annotations if requested
        meta_lines = []
        signer_title = str(sig.get("signer_title") or sig.get("title") or "").strip()
        signer_name = str(sig.get("signer_name") or sig.get("name") or "").strip()
        signer_date = str(sig.get("signer_date") or sig.get("date") or "").strip()
        if signer_title:
            meta_lines.append(signer_title)
        if signer_name:
            meta_lines.append(signer_name)
        if signer_date:
            meta_lines.append(signer_date)

        for pno in target_pnos:
            if not (0 <= pno < total_pages):
                continue

            page = doc[pno]
            vw = page.rect.width
            vh = page.rect.height

            # Determine visual rectangle
            if "x_pct" in sig and "y_pct" in sig:
                x_pct = float(sig.get("x_pct", 0.65))
                y_pct = float(sig.get("y_pct", 0.75))
                w_pct = float(sig.get("w_pct", 0.25))
                h_pct = float(sig.get("h_pct", 0.15))
                v_rect = fitz.Rect(x_pct * vw, y_pct * vh, (x_pct + w_pct) * vw, (y_pct + h_pct) * vh)
            elif "rect" in sig and isinstance(sig["rect"], (list, tuple)) and len(sig["rect"]) == 4:
                r = sig["rect"]
                v_rect = fitz.Rect(float(r[0]), float(r[1]), float(r[2]), float(r[3]))
            else:
                v_rect = fitz.Rect(vw * 0.65, vh * 0.75, vw * 0.95, vh * 0.90)

            # Coordinate transformation for page.rotation
            rot = page.rotation
            if rot != 0:
                target_rect = v_rect * page.derotation_matrix
                insert_rot = rot
            else:
                target_rect = v_rect
                insert_rot = 0

            if img_bytes:
                page.insert_image(target_rect, stream=img_bytes, rotate=insert_rot, keep_proportion=True, overlay=True)
            else:
                page.insert_image(target_rect, filename=img_file, rotate=insert_rot, keep_proportion=True, overlay=True)

            # Insert signer metadata if specified
            if meta_lines:
                text_top = v_rect.y1 + 4
                text_bottom = min(vh - 8, text_top + len(meta_lines) * 14 + 6)
                text_left = max(8.0, v_rect.x0 - 20.0)
                text_right = min(vw - 8.0, v_rect.x1 + 20.0)
                v_text_rect = fitz.Rect(text_left, text_top, text_right, text_bottom)

                target_text_rect = v_text_rect * page.derotation_matrix if rot != 0 else v_text_rect
                fontname = "helv"
                if ttf_font:
                    try:
                        page.insert_font(fontname="sig_meta_font", fontfile=ttf_font)
                        fontname = "sig_meta_font"
                    except Exception:
                        pass
                page.insert_textbox(
                    target_text_rect,
                    "\n".join(meta_lines),
                    fontname=fontname,
                    fontsize=9.0,
                    color=(0.12, 0.16, 0.22),
                    align=fitz.TEXT_ALIGN_CENTER,
                    rotate=insert_rot,
                )

            applied_count += 1

    doc.save(dest, garbage=3, deflate=True)
    doc.close()

    return {
        "success": True,
        "output_path": dest,
        "filename": os.path.basename(dest),
        "total_signatures_applied": applied_count,
        "signatures_applied": applied_count,
    }


def flatten_and_sanitize_pdf(
    pdf_path: str,
    output_path: str,
    flatten_vector: bool = True,
    strict_raster: bool = False,
    sanitize_metadata: bool = True,
    raster_dpi: int = 200,
) -> Dict[str, Any]:
    """
    Hardens a PDF against tampering:
    1. flatten_vector: calls doc.bake(annots=True, widgets=True) to lock forms, fields & stamps into vector stream.
    2. strict_raster: if True, rasterizes each page to high-res images and re-assembles into an un-editable static image PDF.
    3. sanitize_metadata: strips title, author, producer, creation/mod dates, XMP XML, and scrubs attached/embedded files.
    """
    src = os.path.abspath(pdf_path)
    if not os.path.isfile(src):
        raise FileNotFoundError(f"Source PDF not found: {src}")

    dest = os.path.abspath(output_path)
    out_dir = os.path.dirname(dest)
    if out_dir:
        os.makedirs(out_dir, exist_ok=True)

    if strict_raster:
        # Strict Raster Flattening: Render each page to image and rebuild PDF
        src_doc = fitz.open(src)
        new_doc = fitz.open()
        matrix = fitz.Matrix(raster_dpi / 72.0, raster_dpi / 72.0)

        for pno in range(len(src_doc)):
            page = src_doc[pno]
            pw, ph = page.rect.width, page.rect.height
            pix = page.get_pixmap(matrix=matrix, alpha=False)
            img_bytes = pix.tobytes("png")

            new_page = new_doc.new_page(width=pw, height=ph)
            new_page.insert_image(fitz.Rect(0, 0, pw, ph), stream=img_bytes)

        src_doc.close()

        if sanitize_metadata:
            new_doc.set_metadata({})
            try:
                new_doc.del_xml_metadata()
            except Exception:
                pass
            new_doc.scrub(attached_files=True, clean_pages=True, embedded_files=True, metadata=True, xml_metadata=True)

        new_doc.save(dest, garbage=4, deflate=True)
        total_pages = len(new_doc)
        new_doc.close()
        return {
            "success": True,
            "output_path": dest,
            "filename": os.path.basename(dest),
            "mode": "strict_raster",
            "total_pages": total_pages,
            "sanitized": sanitize_metadata,
        }

    # Vector Flattening Mode
    doc = fitz.open(src)

    if flatten_vector:
        try:
            doc.bake(annots=True, widgets=True)
        except Exception as e:
            print(f"[WARN] doc.bake failed, falling back: {e}")

    if sanitize_metadata:
        doc.set_metadata({})
        try:
            doc.del_xml_metadata()
        except Exception:
            pass
        doc.scrub(attached_files=True, clean_pages=True, embedded_files=True, metadata=True, xml_metadata=True)

    doc.save(dest, garbage=4, deflate=True)
    total_pages = len(doc)
    doc.close()

    return {
        "success": True,
        "output_path": dest,
        "filename": os.path.basename(dest),
        "mode": "vector_bake",
        "total_pages": total_pages,
        "sanitized": sanitize_metadata,
    }


def extract_pdf_tables_to_xlsx(
    pdf_path: str,
    output_path: str,
) -> Dict[str, Any]:
    """
    Convenience wrapper around core.converter.table_extractor.extract_pdf_tables_to_xlsx.
    Extracts tabular data from PDF into an enterprise-grade Excel workbook with
    multi-page stitching, ghost column pruning, and 2-tier header reconstruction.
    """
    from core.converter.table_extractor import extract_pdf_tables_to_xlsx as _extract_tables

    ok, res, count = _extract_tables(pdf_path, output_path)
    if not ok:
        return {
            "success": False,
            "error": res,
            "output_path": "",
            "tables_extracted": 0,
        }

    return {
        "success": True,
        "output_path": res,
        "filename": os.path.basename(res),
        "tables_extracted": count,
    }


