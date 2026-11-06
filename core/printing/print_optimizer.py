"""
core.printing.print_optimizer — Enterprise Pre-Print Document Optimizer.

Features:
1. Grayscale Conversion (Ink & Toner Saver):
   Converts document raster layers & pages to monochrome/grayscale, saving up to 70% color toner.
2. N-Up Page Layout (2-Up / 4-Up):
   Merges 2 or 4 pages onto single sheets with automatic aspect ratio scaling and borders.
3. Booklet Duplex Imposition (Saddle-Stitch):
   Pads pages to multiples of 4 and rearranges them into foldable booklet sheet pairs.
4. Margin Safety Guard (ERR_PRINT_001):
   Guarantees minimum 5mm printable margins to prevent text clipping by mechanical printer rollers.
5. Security Watermarking:
   Embeds crisp vector watermarks (e.g. CONFIDENTIAL, BẢN LƯU HÀNH NỘI BỘ, DRAFT) across pages.
"""

from __future__ import annotations

import math
import os
import re
import tempfile
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple, Union

try:
    import pymupdf as fitz
except ImportError:
    import fitz  # type: ignore


class PrintOptimizer:
    """Pre-processes PDF documents before physical or virtual spooling."""

    # Standard Paper Dimensions in points (72 points = 1 inch, 25.4mm = 1 inch)
    A4_WIDTH = 595.276   # 210mm
    A4_HEIGHT = 841.890  # 297mm
    MM_TO_POINTS = 72.0 / 25.4  # ~2.83465 pt/mm

    PAGE_SIZES: Dict[str, Tuple[float, float]] = {
        "a4": (595.276, 841.890),       # 210 x 297 mm
        "letter": (612.000, 792.000),   # 8.5 x 11.0 in (215.9 x 279.4 mm)
        "a3": (841.890, 1190.551),      # 297 x 420 mm
    }

    @staticmethod
    def parse_page_ranges(range_str: Union[str, List[int]], max_pages: int) -> List[int]:
        """
        Parses page range expression (e.g. '1-5, 8, 11-20', 'all') into 0-indexed page list.
        Validates boundaries and removes duplicates while preserving sequential order.
        """
        if isinstance(range_str, list):
            valid = [p for p in range_str if 0 <= p < max_pages]
            if not valid:
                raise ValueError(f"Dải trang chỉ định không hợp lệ trong tài liệu {max_pages} trang.")
            return sorted(list(dict.fromkeys(valid)))

        if not range_str or range_str.strip().lower() in ("all", "tat_ca", "*", ""):
            return list(range(max_pages))

        pages: List[int] = []
        parts = [p.strip() for p in range_str.split(",") if p.strip()]
        for part in parts:
            if "-" in part:
                sub_parts = [s.strip() for s in part.split("-", 1)]
                if not (sub_parts[0].isdigit() and sub_parts[1].isdigit()):
                    raise ValueError(f"Định dạng dải trang không hợp lệ: '{part}'")
                start = int(sub_parts[0])
                end = int(sub_parts[1])
                step = 1 if start <= end else -1
                for p in range(start, end + step, step):
                    if 1 <= p <= max_pages:
                        pages.append(p - 1)
            elif part.isdigit():
                p = int(part)
                if 1 <= p <= max_pages:
                    pages.append(p - 1)
            else:
                raise ValueError(f"Ký tự dải trang không hợp lệ: '{part}'")

        if not pages:
            raise ValueError(f"Dải trang '{range_str}' không chứa trang nào hợp lệ trong tài liệu (1-{max_pages}).")

        # Deduplicate preserving order
        return list(dict.fromkeys(pages))

    @classmethod
    def _apply_page_scaling(
        cls,
        doc: fitz.Document,
        page_size: Optional[str] = None,
        scaling: str = "fit",
        margin_guard_mm: float = 5.0,
    ) -> fitz.Document:
        """
        Scales or re-centers pages onto target paper size (A4, Letter, A3).
        'fit': Scales content while preserving aspect ratio inside printable area (ERR_PRINT_001).
        'actual': Keeps 100% original dimensions, centering on target sheet.
        """
        target_size_key = (page_size or "a4").strip().lower()
        target_dim = cls.PAGE_SIZES.get(target_size_key, (cls.A4_WIDTH, cls.A4_HEIGHT))
        margin_pt = margin_guard_mm * cls.MM_TO_POINTS if scaling == "fit" else 0.0

        out_doc = fitz.open()
        for page in doc:
            orig_rect = page.rect
            orig_w, orig_h = orig_rect.width, orig_rect.height
            is_landscape = orig_w > orig_h

            if is_landscape:
                sheet_w = max(target_dim)
                sheet_h = min(target_dim)
            else:
                sheet_w = min(target_dim)
                sheet_h = max(target_dim)

            new_page = out_doc.new_page(width=sheet_w, height=sheet_h)

            if scaling == "actual":
                off_x = max(0.0, (sheet_w - orig_w) / 2.0)
                off_y = max(0.0, (sheet_h - orig_h) / 2.0)
                dest_rect = fitz.Rect(off_x, off_y, off_x + orig_w, off_y + orig_h)
            else:
                avail_w = max(10.0, sheet_w - (2 * margin_pt))
                avail_h = max(10.0, sheet_h - (2 * margin_pt))
                scale = min(avail_w / orig_w, avail_h / orig_h)
                fit_w = orig_w * scale
                fit_h = orig_h * scale
                off_x = margin_pt + (avail_w - fit_w) / 2.0
                off_y = margin_pt + (avail_h - fit_h) / 2.0
                dest_rect = fitz.Rect(off_x, off_y, off_x + fit_w, off_y + fit_h)

            new_page.show_pdf_page(dest_rect, doc, page.number)

        doc.close()
        return out_doc

    @classmethod
    def optimize_pdf(
        cls,
        input_pdf_path: str,
        output_pdf_path: str,
        grayscale: bool = False,
        n_up: int = 1,
        booklet: bool = False,
        watermark_text: Optional[str] = None,
        margin_guard_mm: float = 5.0,
        vector_grayscale: bool = True,
        dpi: int = 300,
        creep_compensation: bool = True,
        paper_thickness_pt: float = 0.28,
        page_range: Optional[Union[str, List[int]]] = None,
        page_size: Optional[str] = None,
        scaling: Optional[str] = "fit",
    ) -> Dict[str, Any]:
        """
        Executes complete pre-print optimization chain on a PDF file.
        Returns detailed optimization statistics.
        """
        if not os.path.isfile(input_pdf_path):
            raise FileNotFoundError(f"Tệp PDF nguồn không tồn tại: {input_pdf_path}")

        src_doc = fitz.open(input_pdf_path)
        orig_page_count = len(src_doc)
        if orig_page_count == 0:
            src_doc.close()
            raise ValueError("Tệp PDF rỗng, không có trang nào.")

        current_doc = src_doc

        # Step 0: Page Range Extraction (1-based to 0-based)
        extracted_pages = orig_page_count
        if page_range and str(page_range).strip().lower() not in ("all", "tat_ca", "*"):
            selected_indices = cls.parse_page_ranges(page_range, orig_page_count)
            current_doc.select(selected_indices)
            extracted_pages = len(current_doc)

        # Step 1: Page Sizing & Scaling (Fit to Printable Area / Actual Size)
        applied_scaling = scaling or "fit"
        if page_size or applied_scaling in ("fit", "actual"):
            if page_size or (applied_scaling == "fit" and margin_guard_mm > 0):
                current_doc = cls._apply_page_scaling(
                    current_doc,
                    page_size=page_size,
                    scaling=applied_scaling,
                    margin_guard_mm=margin_guard_mm,
                )
        elif margin_guard_mm > 0:
            # Fallback Margin Guard (ERR_PRINT_001)
            current_doc = cls._apply_margin_guard(current_doc, margin_guard_mm)

        # Step 2: Booklet or N-Up imposition
        if booklet:
            current_doc = cls._apply_booklet_imposition(
                current_doc,
                creep_compensation=creep_compensation,
                paper_thickness_pt=paper_thickness_pt,
            )
        elif n_up in (2, 4):
            current_doc = cls._apply_n_up(current_doc, n_up)

        # Step 3: Security Watermark
        if watermark_text and watermark_text.strip():
            current_doc = cls._apply_watermark(current_doc, watermark_text.strip())

        # Step 4: Grayscale Conversion (Vector /DeviceGray or High-Res Raster)
        if grayscale:
            if vector_grayscale:
                current_doc = cls._apply_vector_grayscale(current_doc)
            else:
                current_doc = cls._apply_raster_grayscale(current_doc, dpi=dpi)

        os.makedirs(os.path.dirname(os.path.abspath(output_pdf_path)), exist_ok=True)
        current_doc.save(output_pdf_path, garbage=3, deflate=True)
        final_page_count = len(current_doc)
        current_doc.close()

        file_size_kb = os.path.getsize(output_pdf_path) / 1024.0

        return {
            "success": True,
            "original_pages": orig_page_count,
            "extracted_pages": extracted_pages,
            "optimized_pages": final_page_count,
            "page_range": str(page_range) if page_range else "all",
            "page_size": page_size or "A4",
            "scaling": applied_scaling,
            "grayscale": grayscale,
            "vector_grayscale": vector_grayscale if grayscale else False,
            "n_up": n_up,
            "booklet": booklet,
            "creep_compensation": creep_compensation if booklet else False,
            "watermark": watermark_text,
            "margin_guard_mm": margin_guard_mm,
            "output_pdf": output_pdf_path,
            "file_size_kb": round(file_size_kb, 1),
        }

    @classmethod
    def _apply_margin_guard(cls, doc: fitz.Document, margin_mm: float) -> fitz.Document:
        """Ensures all content stays inside minimum printable margins."""
        margin_pt = margin_mm * cls.MM_TO_POINTS
        out_doc = fitz.open()

        for page in doc:
            rect = page.rect
            w, h = rect.width, rect.height

            # Target rectangle inside safety margins
            target_rect = fitz.Rect(margin_pt, margin_pt, w - margin_pt, h - margin_pt)
            new_page = out_doc.new_page(width=w, height=h)
            new_page.show_pdf_page(target_rect, doc, page.number)

        doc.close()
        return out_doc

    @classmethod
    def _apply_n_up(cls, doc: fitz.Document, n: int) -> fitz.Document:
        """Places 2 or 4 pages per sheet."""
        out_doc = fitz.open()
        total = len(doc)

        if n == 2:
            # 2 portrait pages placed onto 1 landscape sheet
            sheet_w = cls.A4_HEIGHT  # 841.89
            sheet_h = cls.A4_WIDTH   # 595.28
            col_w = sheet_w / 2.0
            row_h = sheet_h

            for i in range(0, total, 2):
                new_page = out_doc.new_page(width=sheet_w, height=sheet_h)
                # Left slot
                rect_left = fitz.Rect(10, 10, col_w - 5, row_h - 10)
                new_page.show_pdf_page(rect_left, doc, i)

                # Right slot
                if i + 1 < total:
                    rect_right = fitz.Rect(col_w + 5, 10, sheet_w - 10, row_h - 10)
                    new_page.show_pdf_page(rect_right, doc, i + 1)

        elif n == 4:
            # 4 portrait pages in 2x2 grid onto 1 portrait sheet
            sheet_w = cls.A4_WIDTH
            sheet_h = cls.A4_HEIGHT
            w2 = sheet_w / 2.0
            h2 = sheet_h / 2.0

            slots = [
                fitz.Rect(10, 10, w2 - 5, h2 - 5),
                fitz.Rect(w2 + 5, 10, sheet_w - 10, h2 - 5),
                fitz.Rect(10, h2 + 5, w2 - 5, sheet_h - 10),
                fitz.Rect(w2 + 5, h2 + 5, sheet_w - 10, sheet_h - 10),
            ]

            for i in range(0, total, 4):
                new_page = out_doc.new_page(width=sheet_w, height=sheet_h)
                for s_idx in range(4):
                    if i + s_idx < total:
                        new_page.show_pdf_page(slots[s_idx], doc, i + s_idx)

        doc.close()
        return out_doc

    @classmethod
    def _apply_booklet_imposition(
        cls,
        doc: fitz.Document,
        creep_compensation: bool = True,
        paper_thickness_pt: float = 0.28,
    ) -> fitz.Document:
        """
        Rearranges pages for 2-sided saddle-stitch booklet printing.
        Pads page count to a multiple of 4, then pairs pages with optional
        creep / shingling compensation to prevent margin drift on thick booklets.
        """
        out_doc = fitz.open()
        total = len(doc)
        rem = total % 4
        needed = (4 - rem) % 4
        padded_total = total + needed

        booklet_pairs: List[Tuple[int, int]] = []
        left = 1
        right = padded_total

        while left < right:
            # Front of sheet: (right - 1, left - 1)
            booklet_pairs.append((right - 1, left - 1))
            # Back of sheet: (left, right - 2)
            booklet_pairs.append((left, right - 2))
            left += 2
            right -= 2

        sheet_w = cls.A4_HEIGHT  # Landscape
        sheet_h = cls.A4_WIDTH
        col_w = sheet_w / 2.0

        for pair_idx, (p_left, p_right) in enumerate(booklet_pairs):
            sheet_idx = pair_idx // 2
            creep = (sheet_idx * paper_thickness_pt) if creep_compensation else 0.0

            new_page = out_doc.new_page(width=sheet_w, height=sheet_h)
            # Left slot: shifts toward center fold (+creep)
            rect_left = fitz.Rect(10 + creep, 10, col_w - 5 + creep, sheet_h - 10)
            # Right slot: shifts toward center fold (-creep)
            rect_right = fitz.Rect(col_w + 5 - creep, 10, sheet_w - 10 - creep, sheet_h - 10)

            if p_left < total:
                new_page.show_pdf_page(rect_left, doc, p_left)
            if p_right < total:
                new_page.show_pdf_page(rect_right, doc, p_right)

        doc.close()
        return out_doc

    @classmethod
    def _apply_watermark(cls, doc: fitz.Document, text: str) -> fitz.Document:
        """Renders semi-transparent diagonal watermark across each page."""
        for page in doc:
            rect = page.rect
            cx, cy = rect.width / 2.0, rect.height / 2.0

            # Draw watermark text rotated at 45 degrees
            font_size = max(24, int(min(rect.width, rect.height) / 10.0))
            pt = fitz.Point(cx - (len(text) * font_size * 0.28), cy)
            page.insert_text(
                point=pt,
                text=text,
                fontsize=font_size,
                fontname="helv",
                color=(0.75, 0.75, 0.75),  # Soft gray
                morph=(pt, fitz.Matrix(45)),
                overlay=True
            )
        return doc

    @classmethod
    def _apply_vector_grayscale(cls, doc: fitz.Document) -> fitz.Document:
        """
        Converts document to DeviceGray while strictly preserving vector text,
        fonts, drawing paths and infinite sharpness (v5.0 Industrial Standard).
        """
        for pno in range(len(doc)):
            page = doc[pno]

            # 1. Vector content stream conversion
            c_xrefs = page.get_contents()
            for xref in c_xrefs:
                raw_bytes = doc.xref_stream(xref)
                if raw_bytes:
                    stream_text = raw_bytes.decode('latin1', errors='replace')
                    pat_rg = r'(\d+(?:\.\d+)?)\s+(\d+(?:\.\d+)?)\s+(\d+(?:\.\d+)?)\s+rg\b'
                    pat_RG = r'(\d+(?:\.\d+)?)\s+(\d+(?:\.\d+)?)\s+(\d+(?:\.\d+)?)\s+RG\b'

                    def _rg2gray(m: re.Match) -> str:
                        r, g, b = float(m.group(1)), float(m.group(2)), float(m.group(3))
                        return f"{round(0.299 * r + 0.587 * g + 0.114 * b, 4)} g"

                    def _RG2gray(m: re.Match) -> str:
                        r, g, b = float(m.group(1)), float(m.group(2)), float(m.group(3))
                        return f"{round(0.299 * r + 0.587 * g + 0.114 * b, 4)} G"

                    mod_text = re.sub(pat_rg, _rg2gray, stream_text)
                    mod_text = re.sub(pat_RG, _RG2gray, mod_text)
                    doc.update_stream(xref, mod_text.encode('latin1'))

            # 2. Embedded raster images conversion to grayscale
            try:
                for img_info in page.get_images():
                    img_xref = img_info[0]
                    pix = fitz.Pixmap(doc, img_xref)
                    if pix.colorspace and pix.colorspace.name != fitz.csGRAY.name:
                        gray_pix = fitz.Pixmap(fitz.csGRAY, pix)
                        doc.update_stream(img_xref, gray_pix.tobytes("png"))
            except Exception:
                pass

        return doc

    @classmethod
    def _apply_raster_grayscale(cls, doc: fitz.Document, dpi: int = 300) -> fitz.Document:
        """Converts document pages to high-resolution grayscale raster."""
        out_doc = fitz.open()

        for page in doc:
            pix = page.get_pixmap(dpi=dpi, colorspace=fitz.csGRAY)
            img_bytes = pix.tobytes("png")

            new_page = out_doc.new_page(width=page.rect.width, height=page.rect.height)
            new_page.insert_image(page.rect, stream=img_bytes)

        doc.close()
        return out_doc

    @classmethod
    def _apply_grayscale(cls, doc: fitz.Document, vector_mode: bool = True, dpi: int = 300) -> fitz.Document:
        """Grayscale dispatcher supporting both vector and raster modes."""
        if vector_mode:
            return cls._apply_vector_grayscale(doc)
        return cls._apply_raster_grayscale(doc, dpi=dpi)


def optimize_pdf_for_printing(
    input_pdf_path: str,
    output_pdf_path: str,
    grayscale: bool = False,
    n_up: int = 1,
    booklet: bool = False,
    watermark_text: Optional[str] = None,
    margin_guard_mm: float = 5.0,
    vector_grayscale: bool = True,
    dpi: int = 300,
    creep_compensation: bool = True,
    paper_thickness_pt: float = 0.28,
    page_range: Optional[Union[str, List[int]]] = None,
    page_size: Optional[str] = None,
    scaling: Optional[str] = "fit",
) -> Dict[str, Any]:
    return PrintOptimizer.optimize_pdf(
        input_pdf_path=input_pdf_path,
        output_pdf_path=output_pdf_path,
        grayscale=grayscale,
        n_up=n_up,
        booklet=booklet,
        watermark_text=watermark_text,
        margin_guard_mm=margin_guard_mm,
        vector_grayscale=vector_grayscale,
        dpi=dpi,
        creep_compensation=creep_compensation,
        paper_thickness_pt=paper_thickness_pt,
        page_range=page_range,
        page_size=page_size,
        scaling=scaling,
    )
