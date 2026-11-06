# -*- coding: utf-8 -*-
"""
tests/test_pdf_form_fields.py — Tests for PDF AcroForms Field Inspector and Filler.
"""

import os
import tempfile
import pytest

try:
    import pymupdf as fitz
    HAS_PYMUPDF = True
except ImportError:
    try:
        import fitz
        HAS_PYMUPDF = True
    except ImportError:
        HAS_PYMUPDF = False

from core.converter.pdf_toolkit import inspect_pdf_form_fields, fill_pdf_form_fields


@pytest.mark.skipif(not HAS_PYMUPDF, reason="PyMuPDF not installed")
def test_inspect_and_fill_pdf_form_fields():
    with tempfile.TemporaryDirectory() as tmp_dir:
        form_pdf = os.path.join(tmp_dir, "form.pdf")
        filled_pdf = os.path.join(tmp_dir, "filled.pdf")

        # Create a test PDF with 2 widgets: text field and checkbox
        doc = fitz.open()
        page = doc.new_page(width=595, height=842)

        w1 = fitz.Widget()
        w1.field_name = "full_name"
        w1.field_type = fitz.PDF_WIDGET_TYPE_TEXT
        w1.rect = fitz.Rect(100, 100, 300, 130)
        w1.field_value = "Initial Name"
        page.add_widget(w1)

        w2 = fitz.Widget()
        w2.field_name = "is_active"
        w2.field_type = fitz.PDF_WIDGET_TYPE_CHECKBOX
        w2.rect = fitz.Rect(100, 150, 130, 180)
        w2.field_value = False
        page.add_widget(w2)

        doc.save(form_pdf)
        doc.close()

        # 1. Test inspect
        inspect_res = inspect_pdf_form_fields(form_pdf)
        assert inspect_res["success"] is True
        assert inspect_res["total_fields"] == 2
        field_names = [f["field_name"] for f in inspect_res["fields"]]
        assert "full_name" in field_names
        assert "is_active" in field_names

        # 2. Test fill
        fill_res = fill_pdf_form_fields(
            form_pdf,
            filled_pdf,
            {"full_name": "Nguyen Van A", "is_active": True},
            flatten=False
        )
        assert fill_res["success"] is True
        assert fill_res["total_filled"] == 2
        assert os.path.isfile(filled_pdf)

        # 3. Verify values in filled PDF
        doc_filled = fitz.open(filled_pdf)
        w_dict = {w.field_name: w.field_value for w in doc_filled[0].widgets()}
        doc_filled.close()

        assert w_dict["full_name"] == "Nguyen Van A"
        assert w_dict["is_active"] in (True, "Yes")


@pytest.mark.skipif(not HAS_PYMUPDF, reason="PyMuPDF not installed")
def test_mcp_pdf_quick_tools_form_actions():
    from server.mcp_server import execute_pdf_quick_tools

    with tempfile.TemporaryDirectory() as tmp_dir:
        form_pdf = os.path.join(tmp_dir, "form.pdf")
        out_pdf = os.path.join(tmp_dir, "out_filled.pdf")

        doc = fitz.open()
        page = doc.new_page(width=595, height=842)
        w = fitz.Widget()
        w.field_name = "email"
        w.field_type = fitz.PDF_WIDGET_TYPE_TEXT
        w.rect = fitz.Rect(50, 50, 200, 80)
        w.field_value = ""
        page.add_widget(w)
        doc.save(form_pdf)
        doc.close()

        # Test inspect via MCP
        insp = execute_pdf_quick_tools({"action": "inspect_form", "input_path": form_pdf})
        assert insp["success"] is True
        assert insp["total_fields"] == 1
        assert insp["fields"][0]["field_name"] == "email"

        # Test fill via MCP
        fill = execute_pdf_quick_tools({
            "action": "fill_form",
            "input_path": form_pdf,
            "output_path": out_pdf,
            "form_data": {"email": "test@domain.vn"},
            "flatten": False
        })
        assert fill["success"] is True
        assert fill["total_filled"] == 1
        assert os.path.isfile(out_pdf)


def test_is_scanned_pdf_detection():
    from core.converter.ocr_pipeline import is_scanned_pdf

    with tempfile.TemporaryDirectory() as tmp_dir:
        digital_pdf = os.path.join(tmp_dir, "digital.pdf")

        doc = fitz.open()
        page = doc.new_page(width=595, height=842)
        # Add lots of digital text
        page.insert_text(fitz.Point(50, 100), "Tài liệu kỹ thuật số có chứa đầy đủ văn bản và đoạn văn bản dài.")
        doc.save(digital_pdf)
        doc.close()

        is_scanned, stats = is_scanned_pdf(digital_pdf)
        assert is_scanned is False
        assert stats["avg_text_length"] > 20

