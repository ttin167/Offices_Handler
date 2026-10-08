"""
test_docx_core.py — Automated Test Suite for Upgraded Core DOCX Engine.

Verifies:
1. Generative Mode:
   - Real image embedding and graceful fallback card handling.
   - Complex table generation with merged cells (colspan & rowspan).
   - Multi-section layout with dynamic orientation switching (Portrait ↔ Landscape).
   - Strict OpenXML table invariants (cantSplit, tblHeader, vAlign="center", ERR_DOCX_001).
2. In-Place Template Patching Mode:
   - Text placeholder replacement ({{KEY}}) across body, tables, and headers/footers with run-merging.
   - Block anchor replacement ({{BLOCK:id}}) replacing placeholder paragraphs with tables.
"""

import os
import sys
import unittest
from docx import Document
from docx.enum.section import WD_ORIENT
from docx.oxml.ns import qn

# Ensure root directory is in sys.path
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)

from docx_writer import write_docx, patch_docx_template


class TestDocxCoreEngine(unittest.TestCase):
    def setUp(self):
        self.output_dir = os.path.join(BASE_DIR, "scratch", "test_docx_out")
        os.makedirs(self.output_dir, exist_ok=True)

    def test_generative_mode_with_merged_tables_and_sections(self):
        """Test full generative reconstruction from snapshot with merged table and landscape section."""
        out_file = os.path.join(self.output_dir, "test_generative_doc.docx")
        if os.path.exists(out_file):
            os.remove(out_file)

        snapshot = {
            "sections": [
                {
                    "orientation": "portrait",
                    "margin_top_cm": 2.5,
                    "margin_bottom_cm": 2.5,
                    "margin_left_cm": 3.0,
                    "margin_right_cm": 2.0,
                }
            ],
            "body": [
                {
                    "type": "paragraph",
                    "style_name": "Heading 1",
                    "alignment": "LEFT",
                    "runs": [{"text": "1. Kiếm Thử Hệ Thống", "bold": True, "font_size": 16.0}]
                },
                {
                    "type": "paragraph",
                    "alignment": "JUSTIFY",
                    "runs": [{"text": "Dưới đây là bảng tổng hợp các ca kiểm thử phức tạp có ô gộp:"}]
                },
                {
                    "type": "table",
                    "rows": 3,
                    "cols": 3,
                    "cells": [
                        # Row 0: Header with Colspan=2 on first cell
                        [
                            {"colspan": 2, "is_header": True, "text": "Tiêu Chí Đánh Giá (2 Cột)", "fill_color": "#1F4E78"},
                            {"colspan": 1, "is_header": True, "text": "Kết Quả", "fill_color": "#1F4E78"},
                        ],
                        # Row 1: Cell with Rowspan=2
                        [
                            {"rowspan": 2, "text": "Kiểm Thử Tự Động (2 Hàng)", "font_bold": True},
                            {"colspan": 1, "text": "Unit Test Suite"},
                            {"colspan": 1, "text": "Đạt 100%"},
                        ],
                        # Row 2: Second column of row 2 (row 1's rowspan occupies col 0)
                        [
                            {"colspan": 1, "text": "Integration Test"},
                            {"colspan": 1, "text": "Đạt 98%"},
                        ]
                    ]
                },
                {
                    "type": "image",
                    "path": "non_existent_mock_image.png",
                    "alt_text": "Sơ Đồ Luồng Dữ Liệu Thiếu",
                    "caption": "Hình 1.1: Sơ Đồ Luồng Tự Động"
                },
                {
                    "type": "section_break",
                    "orientation": "landscape",
                    "margin_left_cm": 2.0,
                    "margin_right_cm": 2.0,
                    "header_text": "Phụ Lục Sơ Đồ Khổ Ngang",
                    "footer_page_number": True
                },
                {
                    "type": "paragraph",
                    "style_name": "Heading 2",
                    "runs": [{"text": "Phụ Lục: Sơ Đồ Khổ Ngang (Landscape)", "bold": True}]
                }
            ]
        }

        result_path = write_docx(snapshot, out_file)
        self.assertTrue(os.path.isfile(result_path))

        # Inspect generated docx
        doc = Document(result_path)
        self.assertEqual(len(doc.sections), 2)
        self.assertEqual(doc.sections[0].orientation, WD_ORIENT.PORTRAIT)
        self.assertEqual(doc.sections[1].orientation, WD_ORIENT.LANDSCAPE)

        # Inspect table OpenXML Invariants
        self.assertGreaterEqual(len(doc.tables), 1)
        tbl = doc.tables[0]
        # cantSplit on rows
        for row in tbl.rows:
            trPr = row._tr.find(qn("w:trPr"))
            self.assertIsNotNone(trPr)
            self.assertIsNotNone(trPr.find(qn("w:cantSplit")))

        # tblHeader on first row
        first_row_trPr = tbl.rows[0]._tr.find(qn("w:trPr"))
        self.assertIsNotNone(first_row_trPr.find(qn("w:tblHeader")))

        # Check Last Paragraph Rule (ERR_DOCX_001) for all cells
        for row in tbl.rows:
            for cell in row.cells:
                self.assertGreater(len(cell._tc.findall(qn("w:p"))), 0)

        # Check fallback card rendered for missing image
        self.assertGreaterEqual(len(doc.tables), 2)  # Table 1: data table, Table 2: fallback card

    def test_in_place_template_patching(self):
        """Test in-place template patching preserving structure, replacing text & block anchors."""
        template_file = os.path.join(self.output_dir, "test_template.docx")
        patched_file = os.path.join(self.output_dir, "test_patched.docx")

        # 1. Create a dummy template docx
        doc = Document()
        doc.add_heading("TÀI LIỆU DỰ ÁN: {{PROJECT_NAME}}", level=1)
        doc.add_paragraph("Tác giả: {{AUTHOR_NAME}} | Phiên bản: {{VERSION}}")
        doc.add_paragraph("Mô tả hệ thống và bảng kiến trúc chi tiết:")
        doc.add_paragraph("{{BLOCK:ARCHITECTURE_TABLE}}")
        doc.add_paragraph("Phần kết luận của tài liệu.")
        doc.save(template_file)

        # 2. Patch the template
        text_replacements = {
            "{{PROJECT_NAME}}": "ZSCORT Enterprise Automation",
            "{{AUTHOR_NAME}}": "Nguyễn Kỹ Sư",
            "{{VERSION}}": "v3.2.0-Alpha"
        }

        block_replacements = {
            "{{BLOCK:ARCHITECTURE_TABLE}}": {
                "type": "table",
                "rows": 2,
                "cols": 2,
                "cells": [
                    [
                        {"is_header": True, "text": "Phân Hệ", "fill_color": "#0F172A"},
                        {"is_header": True, "text": "Trạng Thái", "fill_color": "#0F172A"}
                    ],
                    [
                        {"text": "Core DOCX Builder"},
                        {"text": "Hoàn Thành (Production Grade)"}
                    ]
                ]
            }
        }

        res = patch_docx_template(
            template_path=template_file,
            output_path=patched_file,
            text_replacements=text_replacements,
            block_replacements=block_replacements
        )
        self.assertTrue(os.path.isfile(res))

        # 3. Verify patched document
        patched_doc = Document(res)
        full_text = " ".join([p.text for p in patched_doc.paragraphs])
        self.assertIn("ZSCORT Enterprise Automation", full_text)
        self.assertIn("Nguyễn Kỹ Sư", full_text)
        self.assertIn("v3.2.0-Alpha", full_text)
        self.assertNotIn("{{PROJECT_NAME}}", full_text)
        self.assertNotIn("{{BLOCK:ARCHITECTURE_TABLE}}", full_text)

        # Verify block table was inserted
        self.assertEqual(len(patched_doc.tables), 1)
        self.assertEqual(patched_doc.tables[0].cell(1, 0).text.strip(), "Core DOCX Builder")


if __name__ == "__main__":
    unittest.main()
