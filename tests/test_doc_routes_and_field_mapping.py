# -*- coding: utf-8 -*-
"""
Tests for doc_routes, inspect_data_columns, and field mapping in mail_merge.
"""

import os
import tempfile
import pytest
from docx import Document

from core.docx.mail_merge import inspect_data_columns, execute_mail_merge
from server.routes.base import Router


def test_inspect_data_columns_csv():
    with tempfile.TemporaryDirectory() as tmpdir:
        csv_path = os.path.join(tmpdir, "test.csv")
        with open(csv_path, "w", encoding="utf-8") as f:
            f.write("Full Name,Email,Salary\nNguyen Van A,a@test.com,15000000\nTran Thi B,b@test.com,20000000\n")

        res = inspect_data_columns(csv_path)
        assert res["success"] is True
        assert "Full Name" in res["columns"]
        assert "Email" in res["columns"]
        assert "Salary" in res["columns"]
        assert res["total_records"] == 2
        assert len(res["sample"]) == 2
        assert res["sample"][0]["Full Name"] == "Nguyen Van A"


def test_mail_merge_with_field_mapping():
    with tempfile.TemporaryDirectory() as tmpdir:
        # 1. Create DOCX template with {{fullname}} and {{email}}
        tpl_path = os.path.join(tmpdir, "template.docx")
        doc = Document()
        doc.add_heading("Hello {{fullname}}", level=1)
        doc.add_paragraph("Contact: {{email}}")
        doc.save(tpl_path)

        # 2. Data source with different column headers
        records = [
            {"Ten Nhan Vien": "Hoang Nam", "Dia Chi Email": "nam@example.com"},
            {"Ten Nhan Vien": "Le Huong", "Dia Chi Email": "huong@example.com"},
        ]

        # 3. Field mapping: token -> column
        mapping = {
            "fullname": "Ten Nhan Vien",
            "email": "Dia Chi Email",
        }

        out_dir = os.path.join(tmpdir, "output")
        res = execute_mail_merge(
            template_path=tpl_path,
            data_source=records,
            output_dir=out_dir,
            field_mapping=mapping,
        )

        assert res["success"] is True
        assert res["total_records"] == 2
        assert len(res["generated_files"]) == 2

        # Check content of first generated file
        f1 = res["generated_files"][0]
        d1 = Document(f1)
        full_text = " ".join(p.text for p in d1.paragraphs)
        assert "Hoang Nam" in full_text
        assert "nam@example.com" in full_text
        assert "{{fullname}}" not in full_text


def test_doc_routes_registered():
    router = Router.get_instance()
    assert ("POST", "/api/docx/read") in router._exact_routes
    assert ("POST", "/api/mail-merge/inspect-data") in router._exact_routes

