"""
tests/test_interactive_toc.py — Automated Test Suite for Interactive Clickable Table of Contents.

Verifies:
1. GFM-compliant heading slugification with Vietnamese diacritics & symbols.
2. ISO/IEC 29500 compliant Word bookmark naming (<= 40 chars, [a-zA-Z0-9_]).
3. Embedding of <w:bookmarkStart> and <w:bookmarkEnd> on Word headings.
4. Rendering of <w:hyperlink w:anchor="..."> on TOC dot leaders and Markdown links.
5. End-to-end PDF link annotation preservation (LINK_GOTO internal destinations).
"""

import os
import sys
import unittest
import tempfile
import pymupdf
import docx

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)

from core.converter.markdown import (
    slugify_markdown_heading,
    make_word_bookmark_name,
    generate_markdown_toc,
    markdown_to_docx,
)
from core.converter.engine import convert_universal


class TestInteractiveTOC(unittest.TestCase):
    def test_slugify_heading_vietnamese_and_math(self):
        """Verify slugification handles Vietnamese, math, and punctuation without desync."""
        s1 = slugify_markdown_heading("1. TỔNG QUAN HỆ THỐNG & TRIẾT LÝ THIẾT KẾ 5 TẦNG")
        self.assertEqual(s1, "1-tổng-quan-hệ-thống-triết-lý-thiết-kế-5-tầng")

        s2 = slugify_markdown_heading(r"4.4. Executive Deck Synthesizer: Cho Tài Liệu $\ge 100$ Trang")
        self.assertEqual(s2, "44-executive-deck-synthesizer-cho-tài-liệu-ge-100-trang")

        s3 = slugify_markdown_heading("10. BẢNG MA TRẬN QUY CHUẨN BẤT BIẾN (INVARIANT QUALITY GATES MATRIX)")
        self.assertEqual(s3, "10-bảng-ma-trận-quy-chuẩn-bất-biến-invariant-quality-gates-matrix")

    def test_make_word_bookmark_name_invariants(self):
        """Verify bookmark names strictly comply with Word <= 40 chars and ASCII identifier rules."""
        bm1 = make_word_bookmark_name("1-tổng-quan-hệ-thống-triết-lý-thiết-kế-5-tầng", 1)
        self.assertTrue(len(bm1) <= 40, f"Bookmark length {len(bm1)} > 40")
        self.assertTrue(bm1.startswith("_bm_1_"))
        self.assertTrue(all(c.isalnum() or c == "_" for c in bm1))

        bm_long = make_word_bookmark_name("r" * 80, 99)
        self.assertTrue(len(bm_long) <= 40)
        self.assertTrue(bm_long.startswith("_bm_99_"))

    def test_markdown_toc_generation(self):
        """Verify generate_markdown_toc generates valid nested Markdown list with anchors."""
        sample_md = """# Document Title

## 1. Giới Thiệu
Content 1

### 1.1. Bối Cảnh
Sub content

## 2. Kiến Trúc Hệ Thống
Content 2
"""
        toc = generate_markdown_toc(sample_md, min_level=2, max_level=3)
        self.assertIn("- [1. Giới Thiệu](#1-giới-thiệu)", toc)
        self.assertIn("  - [1.1. Bối Cảnh](#11-bối-cảnh)", toc)
        self.assertIn("- [2. Kiến Trúc Hệ Thống](#2-kiến-trúc-hệ-thống)", toc)

    def test_docx_bookmarks_and_hyperlinks(self):
        """Verify conversion to DOCX embeds bookmarks and hyperlinks properly."""
        sample_md = """# Báo Cáo Thử Nghiệm

## MỤC LỤC TỔNG QUAN

1. [Tổng Quan Hệ Thống](#1-tổng-quan-hệ-thống)
2. [Chi Tiết Kỹ Thuật](#2-chi-tiết-kỹ-thuật)

---

## 1. Tổng Quan Hệ Thống
Nội dung phần 1.

## 2. Chi Tiết Kỹ Thuật
Nội dung phần 2.
"""
        with tempfile.TemporaryDirectory() as tmpdir:
            md_file = os.path.join(tmpdir, "test.md")
            docx_file = os.path.join(tmpdir, "test.docx")
            with open(md_file, "w", encoding="utf-8") as f:
                f.write(sample_md)

            ok, out_docx = markdown_to_docx(md_file, docx_file)
            self.assertTrue(ok)
            self.assertTrue(os.path.isfile(out_docx))

            doc = docx.Document(out_docx)
            xml_str = doc._body._element.xml

            # Check that bookmarkStart elements exist
            self.assertIn("w:bookmarkStart", xml_str)
            self.assertIn("w:bookmarkEnd", xml_str)

            # Check that hyperlink with anchor exists
            self.assertIn("w:hyperlink", xml_str)
            self.assertIn("w:anchor=", xml_str)

    def test_docx_to_pdf_internal_links(self):
        """Verify converting DOCX with bookmarks to PDF creates clickable LINK_GOTO destinations."""
        sample_md = """# Báo Cáo Đích Thân

## MỤC LỤC

- [Phần 1: Giới Thiệu](#phần-1-giới-thiệu)
- [Phần 2: Đánh Giá](#phần-2-đánh-giá)

---

## Phần 1: Giới Thiệu
Đây là nội dung phần 1.

## Phần 2: Đánh Giá
Đây là nội dung phần 2.
"""
        with tempfile.TemporaryDirectory() as tmpdir:
            md_file = os.path.join(tmpdir, "test.md")
            docx_file = os.path.join(tmpdir, "test.docx")
            pdf_file = os.path.join(tmpdir, "test.pdf")
            with open(md_file, "w", encoding="utf-8") as f:
                f.write(sample_md)

            ok1, _, _ = convert_universal(md_file, docx_file)
            self.assertTrue(ok1)
            ok2, _, _ = convert_universal(docx_file, pdf_file)
            self.assertTrue(ok2)
            self.assertTrue(os.path.isfile(pdf_file))

            # Inspect PDF with PyMuPDF
            pdf_doc = pymupdf.open(pdf_file)
            self.assertGreater(len(pdf_doc), 0)

            # Verify that page 1 contains LINK_GOTO links
            page1 = pdf_doc[0]
            links = page1.get_links()
            self.assertGreaterEqual(len(links), 2, "Page 1 must contain at least 2 navigation links")

            # Verify that link kind is 1 (LINK_GOTO destination)
            goto_links = [l for l in links if l.get("kind") == 1]
            self.assertGreaterEqual(len(goto_links), 2, "Links must be internal LINK_GOTO navigation links")
            pdf_doc.close()


if __name__ == "__main__":
    unittest.main()
