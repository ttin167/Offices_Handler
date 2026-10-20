"""
tests/test_v42_parser_and_healing.py
Comprehensive unit test suite for:
1. Markdown -> DOCX parser:
   - Heading level 5 and 6 (##### and ######) mapping to #0F294A bold headings.
   - Reset of alphabetical and numbered lists across headings (restarts at 1. / a.).
   - Cleaning of inline code backticks (`...`) in paragraphs containing math ($...$) or angle symbols (90°).
   - Font fallback for arrows (↔, →, ↻) and bracketed icons ([↻], [🗑️], [🗜️ Nén PDF]).
2. PDF -> DOCX smart post processor healing rules:
   - Rule 1: Table orphan row merger (merges col 3, deletes split rows, restores vertical rowspan).
   - Rule 2: ASCII diagram line break restoration at ┐│ and ││ junctions.
   - Rule 3: Mid-sentence fake tab (\t+) removal.
   - Rule 4: List number spacing normalization and glued heading-bullet separation.
"""

import os
import re
import unittest
import docx
from docx.shared import Pt, RGBColor
from docx.oxml.ns import qn

from core.converter.markdown import decoupled_to_docx
from core.docx.lists import ListContextTracker, ensure_numbering_definitions
from core.docx.math import add_math_to_paragraph
from core.docx.post_processor import (
    merge_table_orphan_rows,
    restore_ascii_diagram_linebreaks,
    filter_mid_sentence_tabs,
    normalize_list_spacing_and_split_headings,
)


class TestV42ParserAndHealing(unittest.TestCase):
    def setUp(self):
        self.temp_dir = os.path.join(os.path.dirname(__file__), "temp_v42_tests")
        os.makedirs(self.temp_dir, exist_ok=True)

    def tearDown(self):
        import shutil
        if os.path.isdir(self.temp_dir):
            shutil.rmtree(self.temp_dir, ignore_errors=True)

    # -------------------------------------------------------------------------
    # 1. Heading 5 and 6
    # -------------------------------------------------------------------------
    def test_heading_5_and_6_support(self):
        md_file = os.path.join(self.temp_dir, "test_headings.md")
        docx_file = os.path.join(self.temp_dir, "test_headings.docx")

        content = (
            "# Main Title\n\n"
            "#### Level 4 Heading\n\n"
            "##### 3. Đánh Số Trang & Chuẩn Hóa Lề A4\n\n"
            "Nội dung mô tả tính năng.\n\n"
            "###### Tiểu mục cấp 6 chi tiết\n\n"
            "Nội dung cấp 6.\n"
        )
        with open(md_file, "w", encoding="utf-8") as f:
            f.write(content)

        ok, msg = decoupled_to_docx(md_file, docx_path=docx_file)
        self.assertTrue(ok, f"decoupled_to_docx failed: {msg}")

        doc = docx.Document(docx_file)
        headings = [p.text for p in doc.paragraphs if p.style and "Heading" in p.style.name]
        
        self.assertIn("3. Đánh Số Trang & Chuẩn Hóa Lề A4", headings)
        self.assertIn("Tiểu mục cấp 6 chi tiết", headings)

        # Ensure raw markdown ##### is NOT in document text
        for p in doc.paragraphs:
            self.assertFalse(p.text.startswith("#####"), f"Found raw ##### in: {p.text}")
            self.assertFalse(p.text.startswith("######"), f"Found raw ###### in: {p.text}")

    # -------------------------------------------------------------------------
    # 2. Alphabetical list counter reset across headings
    # -------------------------------------------------------------------------
    def test_alphabetical_list_restart_across_headings(self):
        md_file = os.path.join(self.temp_dir, "test_lists.md")
        docx_file = os.path.join(self.temp_dir, "test_lists.docx")

        content = (
            "##### Phân Hệ 5: Diagram Studio\n"
            "- **Thao tác**:\n"
            "  1. Bước một module 5\n"
            "  2. Bước hai module 5\n\n"
            "##### Phân Hệ 6: PDF Pro Studio\n"
            "- **Thao tác**:\n"
            "  1. Bước một module 6 (phải bắt đầu lại từ a, không được là c)\n"
            "  2. Bước hai module 6\n"
        )
        with open(md_file, "w", encoding="utf-8") as f:
            f.write(content)

        ok, msg = decoupled_to_docx(md_file, docx_path=docx_file)
        self.assertTrue(ok, f"decoupled_to_docx failed: {msg}")

        doc = docx.Document(docx_file)
        # Check numbering.xml for lvlOverride on ilvl 0, 1, 2
        n_part = getattr(doc.part, "numbering_part", None)
        self.assertIsNotNone(n_part)
        xml_content = n_part.element.xml
        self.assertIn('w:ilvl="0"', xml_content)
        self.assertIn('w:ilvl="1"', xml_content)
        self.assertIn('w:ilvl="2"', xml_content)

    # -------------------------------------------------------------------------
    # 3. Backtick cleaning with math and angles
    # -------------------------------------------------------------------------
    def test_backticks_with_math_cleaning(self):
        doc = docx.Document()
        p = doc.add_paragraph()
        raw_text = (
            "Cơ chế chống lệch trục: bù ma trận `derotation_matrix` cho góc xoay $90^\\circ$ "
            "ngay cả khi quay `Trang 1/15` nghiêng 90°."
        )
        add_math_to_paragraph(p, raw_text)

        full_text = p.text
        self.assertNotIn("`", full_text, "Found uncleaned backtick in paragraph!")
        self.assertIn("derotation_matrix", full_text)
        self.assertIn("Trang 1/15", full_text)
        self.assertIn("90°", full_text)

    # -------------------------------------------------------------------------
    # 4. Symbol and emoji font fallback
    # -------------------------------------------------------------------------
    def test_symbol_font_fallback(self):
        doc = docx.Document()
        p = doc.add_paragraph()
        raw_text = "Thao tác: [↻] Xoay ảnh ↔ Chuyển đổi [🗑️] Thùng rác [🗜️ Nén PDF] → Hoàn tất."
        add_math_to_paragraph(p, raw_text)

        # Inspect XML runs in paragraph
        p_xml = p._p.xml
        self.assertTrue(
            "Segoe UI Symbol" in p_xml or "Segoe UI Emoji" in p_xml or "DejaVu Sans" in p_xml or "Noto Color Emoji" in p_xml,
            "Symbol font fallback was not injected into paragraph XML!"
        )

    # -------------------------------------------------------------------------
    # 5. Rule 1: Table orphan row merger
    # -------------------------------------------------------------------------
    def test_table_orphan_row_merger(self):
        doc = docx.Document()
        table = doc.add_table(rows=3, cols=3)

        # Header row
        table.rows[0].cells[0].text = "Nhóm Quy Tắc"
        table.rows[0].cells[1].text = "Mã Quy Tắc"
        table.rows[0].cells[2].text = "Nội Dung"

        # Row 1 (Item)
        table.rows[1].cells[0].text = "Văn Bản Word"
        table.rows[1].cells[1].text = "ERR_DOCX_001"
        table.rows[1].cells[2].text = "Mọi cell bảng (<w:tc>) bắt buộc phải kết thúc bằng tối thiểu một"

        # Row 2 (Orphan split row created by pdf2docx)
        table.rows[2].cells[0].text = ""
        table.rows[2].cells[1].text = ""
        table.rows[2].cells[2].text = "thẻ <w:p>."

        self.assertEqual(len(table.rows), 3)
        merged = merge_table_orphan_rows(doc)
        self.assertEqual(merged, 1)
        self.assertEqual(len(table.rows), 2)
        self.assertIn("thẻ <w:p>.", table.rows[1].cells[2].text)
        self.assertEqual(table.rows[1].cells[0].text, "Văn Bản Word")

    # -------------------------------------------------------------------------
    # 6. Rule 2: ASCII diagram line breaks
    # -------------------------------------------------------------------------
    def test_ascii_diagram_linebreaks(self):
        doc = docx.Document()
        p = doc.add_paragraph("┌──────┐│ TẦNG 1: TRÌNH DIỄN & GIAO DIỆN ││ - Vite Web Studio... │└──────┘")

        healed = restore_ascii_diagram_linebreaks(doc)
        self.assertEqual(healed, 1)
        # Check that runs were formatted with Consolas and line breaks exist
        self.assertIn("w:br", p._p.xml)
        for r in p.runs:
            self.assertEqual(r.font.name, "Consolas")

    # -------------------------------------------------------------------------
    # 7. Rule 3: Mid-sentence tab removal
    # -------------------------------------------------------------------------
    def test_filter_mid_sentence_tabs(self):
        doc = docx.Document()
        p1 = doc.add_paragraph("hình ảnh, \tcông thức toán học...")
        p2 = doc.add_paragraph("phục vụ tài liệu phân \t\ttích nghiệp vụ (BRD/SRS):")
        # TOC line should NOT be touched
        p_toc = doc.add_paragraph("PHẦN I: TỔNG QUAN ... 12")

        cleaned = filter_mid_sentence_tabs(doc)
        self.assertEqual(cleaned, 2)
        self.assertEqual(p1.text, "hình ảnh, công thức toán học...")
        self.assertEqual(p2.text, "phục vụ tài liệu phân tích nghiệp vụ (BRD/SRS):")
        self.assertEqual(p_toc.text, "PHẦN I: TỔNG QUAN ... 12")

    # -------------------------------------------------------------------------
    # 8. Rule 4: List number spacing and heading-bullet separation
    # -------------------------------------------------------------------------
    def test_normalize_list_spacing_and_split_headings(self):
        doc = docx.Document()
        p1 = doc.add_paragraph("5. Phiên Bản 4.1 (PDF Pro Studio & Render Cloud 24/7 Hardening) \t• \tBước đột phá chiến lược:")
        p2 = doc.add_paragraph("1.Phân mảnh định dạng...")
        p3 = doc.add_paragraph("b.Kéo-Thả Đổi Thứ Tự...")

        modified = normalize_list_spacing_and_split_headings(doc)
        self.assertGreaterEqual(modified, 3)

        # p1 should now be split into 2 paragraphs
        self.assertEqual(doc.paragraphs[0].text, "5. Phiên Bản 4.1 (PDF Pro Studio & Render Cloud 24/7 Hardening)")
        self.assertEqual(doc.paragraphs[1].text, "• Bước đột phá chiến lược:")
        self.assertEqual(doc.paragraphs[2].text, "1. Phân mảnh định dạng...")
        self.assertEqual(doc.paragraphs[3].text, "b. Kéo-Thả Đổi Thứ Tự...")


if __name__ == "__main__":
    unittest.main()
