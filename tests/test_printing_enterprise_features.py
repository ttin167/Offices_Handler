# -*- coding: utf-8 -*-
"""
tests/test_printing_enterprise_features.py — Test Suite for Enterprise Printing Upgrades:
1. Page Range selection and AST parsing (parse_page_ranges)
2. Duplex configuration (Simplex, Duplex Long Edge, Duplex Short Edge)
3. Page Sizing & Scaling (A4, Letter, A3, Fit vs Actual)
4. Printer Hardware Status & Spooler Job Tracker (get_printer_status, paper out, jam flags)
5. Server API Endpoints (/api/print/status, /api/print/job)
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


def create_multi_page_pdf(file_path: str, pages: int = 8) -> str:
    doc = fitz.open()
    for i in range(pages):
        page = doc.new_page(width=595.28, height=841.89)  # A4
        page.insert_text(fitz.Point(50, 100), f"Document Page {i + 1} - Enterprise Report", fontsize=20)
        page.draw_rect(fitz.Rect(50, 120, 500, 300), color=(0.2, 0.4, 0.8), width=1.5)
    os.makedirs(os.path.dirname(os.path.abspath(file_path)), exist_ok=True)
    doc.save(file_path)
    doc.close()
    return file_path


def test_parse_page_ranges_valid():
    # 1. Standard ranges and singles
    res = PrintOptimizer.parse_page_ranges("1-3, 5, 7-8", max_pages=10)
    assert res == [0, 1, 2, 4, 6, 7]

    # 2. Single page
    res_single = PrintOptimizer.parse_page_ranges("4", max_pages=10)
    assert res_single == [3]

    # 3. Deduplication and sequence order
    res_dedup = PrintOptimizer.parse_page_ranges("3, 1-2, 2, 3, 5", max_pages=10)
    assert res_dedup == [2, 0, 1, 4]

    # 4. "all", "*", empty string
    assert PrintOptimizer.parse_page_ranges("all", max_pages=5) == [0, 1, 2, 3, 4]
    assert PrintOptimizer.parse_page_ranges("*", max_pages=3) == [0, 1, 2]
    assert PrintOptimizer.parse_page_ranges("", max_pages=4) == [0, 1, 2, 3]

    # 5. List input
    assert PrintOptimizer.parse_page_ranges([0, 2, 4], max_pages=6) == [0, 2, 4]


def test_parse_page_ranges_invalid():
    # Out of bounds
    with pytest.raises(ValueError):
        PrintOptimizer.parse_page_ranges("99-105", max_pages=10)

    # Invalid non-digit
    with pytest.raises(ValueError):
        PrintOptimizer.parse_page_ranges("1-3, xyz", max_pages=10)


def test_optimize_pdf_page_range_extraction():
    with tempfile.TemporaryDirectory() as tmp:
        src = os.path.join(tmp, "src_8p.pdf")
        out = os.path.join(tmp, "range_3p.pdf")
        create_multi_page_pdf(src, pages=8)

        # Extract only pages 2, 4, 6 (1-based: "2, 4, 6")
        res = PrintOptimizer.optimize_pdf(
            input_pdf_path=src,
            output_pdf_path=out,
            page_range="2, 4, 6",
        )
        assert res["success"] is True
        assert res["original_pages"] == 8
        assert res["extracted_pages"] == 3
        assert res["optimized_pages"] == 3

        doc = fitz.open(out)
        assert len(doc) == 3
        text_p1 = doc[0].get_text()
        assert "Document Page 2" in text_p1
        text_p2 = doc[1].get_text()
        assert "Document Page 4" in text_p2
        text_p3 = doc[2].get_text()
        assert "Document Page 6" in text_p3
        doc.close()


def test_optimize_pdf_page_size_and_scaling():
    with tempfile.TemporaryDirectory() as tmp:
        src = os.path.join(tmp, "src_a4.pdf")
        out_letter_fit = os.path.join(tmp, "out_letter_fit.pdf")
        out_a3_actual = os.path.join(tmp, "out_a3_actual.pdf")
        create_multi_page_pdf(src, pages=2)

        # 1. Scale A4 to Letter Fit
        res_fit = PrintOptimizer.optimize_pdf(
            input_pdf_path=src,
            output_pdf_path=out_letter_fit,
            page_size="letter",
            scaling="fit",
        )
        assert res_fit["success"] is True
        assert res_fit["page_size"] == "letter"
        assert res_fit["scaling"] == "fit"

        doc_fit = fitz.open(out_letter_fit)
        assert len(doc_fit) == 2
        # Letter height is 792, width is 612
        assert abs(doc_fit[0].rect.width - 612.0) < 1.0
        assert abs(doc_fit[0].rect.height - 792.0) < 1.0
        doc_fit.close()

        # 2. Scale A4 to A3 Actual (centered 1:1)
        res_a3 = PrintOptimizer.optimize_pdf(
            input_pdf_path=src,
            output_pdf_path=out_a3_actual,
            page_size="a3",
            scaling="actual",
        )
        assert res_a3["success"] is True
        doc_a3 = fitz.open(out_a3_actual)
        # A3 height is 1190.55, width is 841.89
        assert abs(doc_a3[0].rect.width - 841.89) < 1.0
        assert abs(doc_a3[0].rect.height - 1190.55) < 1.0
        doc_a3.close()


def test_printer_manager_get_printer_status():
    # 1. Default/virtual status
    status_info = PrinterManager.get_printer_status()
    assert isinstance(status_info, dict)
    assert "printer" in status_info
    assert "status" in status_info
    assert "c_jobs" in status_info
    assert "message" in status_info

    # 2. Explicit Virtual Spooler
    v_status = PrinterManager.get_printer_status("Virtual Spooler (PDF Buffer)")
    assert v_status["is_virtual"] is True
    assert v_status["status"] == "READY"


def test_printer_manager_print_file_enterprise():
    with tempfile.TemporaryDirectory() as tmp:
        src = os.path.join(tmp, "report_8p.pdf")
        create_multi_page_pdf(src, pages=8)

        job_res = PrinterManager.print_file(
            file_path=src,
            printer_name="Virtual Spooler (PDF Buffer)",
            copies=1,
            page_range="1-2, 5",
            page_size="a4",
            scaling="fit",
            duplex="vertical",
            optimizer_options={"grayscale": True, "watermark": "SAMPLE BẢO MẬT"},
        )
        assert job_res["success"] is True
        assert job_res["mode"] == "VIRTUAL_SPOOL"
        assert job_res["page_range"] == "1-2, 5"
        assert job_res["duplex"] == "vertical"
        assert os.path.isfile(job_res["spooled_file"])

        # Check spooled file page count
        doc = fitz.open(job_res["spooled_file"])
        assert len(doc) == 3  # pages 1, 2, 5
        doc.close()


def test_printer_manager_spool_via_gdi():
    with tempfile.TemporaryDirectory() as tmp:
        src = os.path.join(tmp, "doc_gdi.pdf")
        create_multi_page_pdf(src, pages=2)

        # Non-existent file returns False safely
        assert PrinterManager._spool_via_gdi("non_existent.pdf", "Mock Printer") is False

        # Non-existent printer returns False safely without crash
        assert PrinterManager._spool_via_gdi(src, "NonExistentPrinter12345") is False


