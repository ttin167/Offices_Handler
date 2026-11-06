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
import tempfile
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple, Union

try:
    import pymupdf as fitz
except ImportError:
    import fitz  # type: ignore


class PrintOptimizer:
    """Pre-processes PDF documents before physical or virtual spooling."""

    # Standard A4 dimensions in points (72 points = 1 inch, 25.4mm = 1 inch)
    A4_WIDTH = 595.276   # 210mm
    A4_HEIGHT = 841.890  # 297mm
    MM_TO_POINTS = 72.0 / 25.4  # ~2.83465 pt/mm

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

        # Step 1: Margin Guard (ERR_PRINT_001) - Shrink if content is too close to border
        if margin_guard_mm > 0:
            current_doc = cls._apply_margin_guard(current_doc, margin_guard_mm)

        # Step 2: Booklet or N-Up imposition
        if booklet:
            current_doc = cls._apply_booklet_imposition(current_doc)
        elif n_up in (2, 4):
            current_doc = cls._apply_n_up(current_doc, n_up)

        # Step 3: Security Watermark
        if watermark_text and watermark_text.strip():
            current_doc = cls._apply_watermark(current_doc, watermark_text.strip())

        # Step 4: Grayscale Conversion
        if grayscale:
            current_doc = cls._apply_grayscale(current_doc)

        os.makedirs(os.path.dirname(os.path.abspath(output_pdf_path)), exist_ok=True)
        current_doc.save(output_pdf_path, garbage=3, deflate=True)
        final_page_count = len(current_doc)
        current_doc.close()

        file_size_kb = os.path.getsize(output_pdf_path) / 1024.0

        return {
            "success": True,
            "original_pages": orig_page_count,
            "optimized_pages": final_page_count,
            "grayscale": grayscale,
            "n_up": n_up,
            "booklet": booklet,
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
    def _apply_booklet_imposition(cls, doc: fitz.Document) -> fitz.Document:
        """
        Rearranges pages for 2-sided saddle-stitch booklet printing.
        Pads page count to a multiple of 4, then pairs pages.
        """
        out_doc = fitz.open()
        total = len(doc)
        rem = total % 4
        needed = (4 - rem) % 4
        padded_total = total + needed

        # Virtual page order for booklet:
        # Sheet 1 Front: (Page N, Page 1)
        # Sheet 1 Back:  (Page 2, Page N-1)
        # Sheet 2 Front: (Page N-2, Page 3)
        # etc.
        booklet_pairs: List[Tuple[int, int]] = []
        left = 1
        right = padded_total

        while left < right:
            # Front of sheet: (right, left)
            booklet_pairs.append((right - 1, left - 1))
            # Back of sheet: (left + 1, right - 1)
            booklet_pairs.append((left, right - 2))
            left += 2
            right -= 2

        sheet_w = cls.A4_HEIGHT  # Landscape
        sheet_h = cls.A4_WIDTH
        col_w = sheet_w / 2.0

        for p_left, p_right in booklet_pairs:
            new_page = out_doc.new_page(width=sheet_w, height=sheet_h)
            rect_left = fitz.Rect(10, 10, col_w - 5, sheet_h - 10)
            rect_right = fitz.Rect(col_w + 5, 10, sheet_w - 10, sheet_h - 10)

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
    def _apply_grayscale(cls, doc: fitz.Document) -> fitz.Document:
        """Converts document pages to pure grayscale raster to save color toner."""
        out_doc = fitz.open()

        for page in doc:
            # Render page to grayscale pixmap at 200 DPI for print sharpness
            pix = page.get_pixmap(dpi=200, colorspace=fitz.csGRAY)
            img_bytes = pix.tobytes("png")

            new_page = out_doc.new_page(width=page.rect.width, height=page.rect.height)
            new_page.insert_image(page.rect, stream=img_bytes)

        doc.close()
        return out_doc


def optimize_pdf_for_printing(
    input_pdf_path: str,
    output_pdf_path: str,
    grayscale: bool = False,
    n_up: int = 1,
    booklet: bool = False,
    watermark_text: Optional[str] = None,
    margin_guard_mm: float = 5.0,
) -> Dict[str, Any]:
    return PrintOptimizer.optimize_pdf(
        input_pdf_path=input_pdf_path,
        output_pdf_path=output_pdf_path,
        grayscale=grayscale,
        n_up=n_up,
        booklet=booklet,
        watermark_text=watermark_text,
        margin_guard_mm=margin_guard_mm,
    )
