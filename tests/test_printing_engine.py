"""
tests/test_printing_engine.py — Test Suite for Enterprise Physical & Virtual Printing Engine.
"""

import os
import tempfile
import pytest

try:
    import fitz
except ImportError:
    import pymupdf as fitz

from core.printing.print_optimizer import PrintOptimizer
from core.printing.printer_manager import PrinterManager


def create_sample_pdf(file_path: str, pages: int = 4) -> str:
    doc = fitz.open()
    for i in range(pages):
        page = doc.new_page(width=595.28, height=841.89)  # A4
        page.insert_text(fitz.Point(100, 100), f"Page {i + 1} Content for Printing Test", fontsize=18)
    os.makedirs(os.path.dirname(os.path.abspath(file_path)), exist_ok=True)
    doc.save(file_path)
    doc.close()
    return file_path


def test_printer_manager_discovery():
    printers = PrinterManager.list_printers()
    assert len(printers) > 0
    assert any("name" in p for p in printers)
    # Ensure default printer is string or None
    def_p = PrinterManager.get_default_printer()
    assert def_p is None or isinstance(def_p, str)


def test_print_optimizer_grayscale_and_watermark():
    with tempfile.TemporaryDirectory() as tmp:
        src_pdf = os.path.join(tmp, "source.pdf")
        opt_pdf = os.path.join(tmp, "optimized.pdf")
        create_sample_pdf(src_pdf, pages=2)

        res = PrintOptimizer.optimize_pdf(
            input_pdf_path=src_pdf,
            output_pdf_path=opt_pdf,
            grayscale=True,
            watermark_text="BẢN LƯU NỘI BỘ",
            n_up=1
        )
        assert res["success"] is True
        assert res["original_pages"] == 2
        assert res["optimized_pages"] == 2
        assert os.path.isfile(opt_pdf)


def test_print_optimizer_nup_and_booklet():
    with tempfile.TemporaryDirectory() as tmp:
        src_pdf = os.path.join(tmp, "source_4p.pdf")
        nup_pdf = os.path.join(tmp, "nup_2.pdf")
        booklet_pdf = os.path.join(tmp, "booklet.pdf")
        create_sample_pdf(src_pdf, pages=4)

        # 2-Up
        res_2up = PrintOptimizer.optimize_pdf(src_pdf, nup_pdf, n_up=2)
        assert res_2up["success"] is True
        assert res_2up["optimized_pages"] == 2

        # Booklet
        res_bkl = PrintOptimizer.optimize_pdf(src_pdf, booklet_pdf, booklet=True)
        assert res_bkl["success"] is True
        assert res_bkl["optimized_pages"] == 2


def test_printer_manager_virtual_spool():
    with tempfile.TemporaryDirectory() as tmp:
        src_pdf = os.path.join(tmp, "job_doc.pdf")
        create_sample_pdf(src_pdf, pages=1)

        job_res = PrinterManager.print_file(
            file_path=src_pdf,
            printer_name="Virtual Spooler (PDF Buffer)",
            copies=1,
            optimizer_options={"grayscale": True, "watermark": "SAMPLE"}
        )
        assert job_res["success"] is True
        assert job_res["mode"] == "VIRTUAL_SPOOL"
        assert os.path.isfile(job_res["spooled_file"])

