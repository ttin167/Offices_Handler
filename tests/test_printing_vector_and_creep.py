# -*- coding: utf-8 -*-
"""
tests/test_printing_vector_and_creep.py — Test Suite for Industrial Printing Enhancements.
"""

import os
import tempfile
import pytest

try:
    import pymupdf as fitz
except ImportError:
    import fitz  # type: ignore

from core.printing.print_optimizer import PrintOptimizer
from core.printing.printer_manager import PrinterManager


def create_sample_color_pdf(path: str, pages: int = 8) -> str:
    doc = fitz.open()
    for i in range(pages):
        p = doc.new_page(width=595.28, height=841.89)
        # Vector text with RGB color
        p.insert_text(fitz.Point(100, 100), f"Page {i + 1} Vector Text", fontsize=24, color=(1, 0, 0))
        # Vector rectangle with RGB stroke and fill
        p.draw_rect(fitz.Rect(100, 150, 400, 250), color=(0, 0, 1), fill=(0, 0.8, 0.2))
    os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
    doc.save(path)
    doc.close()
    return path


def test_vector_grayscale_preservation():
    with tempfile.TemporaryDirectory() as tmp:
        src_pdf = os.path.join(tmp, "src_vector.pdf")
        opt_pdf = os.path.join(tmp, "opt_vector.pdf")
        create_sample_color_pdf(src_pdf, pages=2)

        res = PrintOptimizer.optimize_pdf(
            input_pdf_path=src_pdf,
            output_pdf_path=opt_pdf,
            grayscale=True,
            vector_grayscale=True,
        )
        assert res["success"] is True
        assert res["vector_grayscale"] is True
        assert os.path.isfile(opt_pdf)

        # Verify that text is still vector text and was NOT rasterized
        doc = fitz.open(opt_pdf)
        assert len(doc) == 2
        text_p1 = doc[0].get_text()
        assert "Page 1 Vector Text" in text_p1
        # Embedded images should remain 0 because it was kept as vector!
        assert len(doc[0].get_images()) == 0
        doc.close()


def test_booklet_creep_shingling():
    with tempfile.TemporaryDirectory() as tmp:
        src_pdf = os.path.join(tmp, "src_booklet.pdf")
        opt_pdf = os.path.join(tmp, "opt_booklet.pdf")
        create_sample_color_pdf(src_pdf, pages=8)

        res = PrintOptimizer.optimize_pdf(
            input_pdf_path=src_pdf,
            output_pdf_path=opt_pdf,
            booklet=True,
            creep_compensation=True,
            paper_thickness_pt=0.5,
        )
        assert res["success"] is True
        assert res["booklet"] is True
        assert res["creep_compensation"] is True
        assert res["optimized_pages"] == 4  # 8 pages -> 4 landscape sheets (2-up pairs)
        assert os.path.isfile(opt_pdf)


def test_guard_g2_device_capabilities():
    # Test on default or virtual printer
    caps = PrinterManager.get_device_capabilities()
    assert isinstance(caps, dict)
    assert "supported_bins" in caps
    assert "supports_duplex" in caps

    # Safe apply_devmode with invalid tray should not raise exception (Guard G2)
    applied = PrinterManager.apply_devmode(
        printer_name=caps.get("printer") or "Virtual Spooler",
        tray_id=99999,  # Non-existent tray
        duplex="vertical",
    )
    # Should safely return False without exception
    assert applied is False

