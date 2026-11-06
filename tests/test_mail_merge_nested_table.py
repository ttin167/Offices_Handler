# -*- coding: utf-8 -*-
"""
tests/test_mail_merge_nested_table.py — Tests for Nested Table Block Repeater in Mail Merge.
"""

import os
import tempfile
import docx
from docx import Document
from core.docx.mail_merge import execute_mail_merge


def test_mail_merge_nested_table_repeater():
    with tempfile.TemporaryDirectory() as tmp_dir:
        tpl_path = os.path.join(tmp_dir, "invoice_template.docx")
        out_dir = os.path.join(tmp_dir, "output")

        # Create template with header info and a dynamic table
        doc = Document()
        doc.add_paragraph("Kính gửi khách hàng: {{customer_name}}")

        table = doc.add_table(rows=2, cols=3)
        # Header row
        table.rows[0].cells[0].text = "STT"
        table.rows[0].cells[1].text = "Mặt hàng"
        table.rows[0].cells[2].text = "Đơn giá"

        # Template repeating row with {{#TABLE:items}}
        table.rows[1].cells[0].text = "{{#TABLE:items}}{{stt}}"
        table.rows[1].cells[1].text = "{{item.name}}"
        table.rows[1].cells[2].text = "{{item.price}}{{/TABLE:items}}"

        doc.save(tpl_path)

        # Dataset with 1 invoice having 3 line items
        data = [
            {
                "customer_name": "Cong Ty ABC",
                "items": [
                    {"name": "Dich vu Cloud Server", "price": "1,500,000 VND"},
                    {"name": "Bang thong 1Gbps", "price": "500,000 VND"},
                    {"name": "Sao luu du phong", "price": "300,000 VND"},
                ]
            }
        ]

        res = execute_mail_merge(
            template_path=tpl_path,
            data_source=data,
            output_dir=out_dir,
            filename_pattern="invoice_{{customer_name}}.docx"
        )

        assert res["success"] is True
        assert res["total_records"] == 1
        generated_file = res["generated_files"][0]
        assert os.path.isfile(generated_file)

        # Inspect generated docx
        gen_doc = Document(generated_file)
        assert "Cong Ty ABC" in gen_doc.paragraphs[0].text

        gen_table = gen_doc.tables[0]
        # Should now have 1 header row + 3 expanded items = 4 rows!
        assert len(gen_table.rows) == 4

        # Row 1: STT 1, Dich vu Cloud Server, 1,500,000 VND
        assert gen_table.rows[1].cells[0].text.strip() == "1"
        assert "Cloud Server" in gen_table.rows[1].cells[1].text
        assert "1,500,000 VND" in gen_table.rows[1].cells[2].text

        # Row 3: STT 3, Sao luu du phong, 300,000 VND
        assert gen_table.rows[3].cells[0].text.strip() == "3"
        assert "Sao luu" in gen_table.rows[3].cells[1].text

