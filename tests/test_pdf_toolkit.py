"""
test_pdf_toolkit.py — Unit Tests for Everyday PDF Toolkit (v4.0).

Verifies:
1. merge_pdfs: Concatenates multiple PDFs while preserving page integrity.
2. split_pdf: Splits PDF into individual single pages and custom page ranges.
3. rotate_and_clean_pdf: Rotates pages clockwise and detects/removes blank pages.
4. stamp_watermark_pdf: Injects semi-transparent diagonal watermark text across all pages.
5. compress_pdf: Performs deflation, stream compression, and xref garbage collection.
"""

import os
import sys
import unittest

try:
    import pymupdf as fitz
except ImportError:
    import fitz

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)

from pdf_toolkit import (
    merge_pdfs,
    split_pdf,
    rotate_and_clean_pdf,
    stamp_watermark_pdf,
    compress_pdf,
)


class TestPDFToolkit(unittest.TestCase):
    def setUp(self):
        self.scratch_dir = os.path.join(BASE_DIR, "scratch", "test_pdf_toolkit")
        os.makedirs(self.scratch_dir, exist_ok=True)

    def _create_sample_pdf(self, path: str, num_pages: int, with_blanks: bool = False) -> str:
        doc = fitz.open()
        for i in range(num_pages):
            page = doc.new_page(width=595, height=842)  # A4 pt
            # If with_blanks is True, make page 2 (index 1) blank
            if with_blanks and i == 1:
                continue
            page.insert_text(fitz.Point(100, 150), f"Document Page {i + 1} Content")
        doc.save(path)
        doc.close()
        return path

    def test_merge_pdfs(self):
        """Test concatenating multiple PDFs in sequential order."""
        pdf1 = os.path.join(self.scratch_dir, "doc1.pdf")
        pdf2 = os.path.join(self.scratch_dir, "doc2.pdf")
        merged = os.path.join(self.scratch_dir, "merged.pdf")

        self._create_sample_pdf(pdf1, 2)
        self._create_sample_pdf(pdf2, 3)

        out_path = merge_pdfs([pdf1, pdf2], merged)
        self.assertTrue(os.path.isfile(out_path))

        doc = fitz.open(out_path)
        self.assertEqual(len(doc), 5)
        self.assertIn("Page 1", doc[0].get_text())
        self.assertIn("Page 2", doc[1].get_text())
        doc.close()

    def test_split_pdf_individual_and_ranges(self):
        """Test splitting PDF into individual pages and multi-page range groups."""
        src_pdf = os.path.join(self.scratch_dir, "multi_page.pdf")
        self._create_sample_pdf(src_pdf, 5)

        # 1. Split all individual pages
        split_dir_all = os.path.join(self.scratch_dir, "split_all")
        files_all = split_pdf(src_pdf, split_dir_all)
        self.assertEqual(len(files_all), 5)
        for f in files_all:
            self.assertTrue(os.path.isfile(f))
            sub_doc = fitz.open(f)
            self.assertEqual(len(sub_doc), 1)
            sub_doc.close()

        # 2. Split with ranges: '1-2, 3-5'
        split_dir_ranges = os.path.join(self.scratch_dir, "split_ranges")
        files_ranges = split_pdf(src_pdf, split_dir_ranges, page_ranges="1-2, 3-5")
        self.assertEqual(len(files_ranges), 2)

        part1 = fitz.open(files_ranges[0])
        self.assertEqual(len(part1), 2)
        part1.close()

        part2 = fitz.open(files_ranges[1])
        self.assertEqual(len(part2), 3)
        part2.close()

    def test_rotate_and_clean_pdf(self):
        """Test rotating pages and removing blank pages."""
        dirty_pdf = os.path.join(self.scratch_dir, "dirty_with_blank.pdf")
        cleaned_pdf = os.path.join(self.scratch_dir, "cleaned_rotated.pdf")

        # 3 pages: Page 1 (text), Page 2 (blank), Page 3 (text)
        self._create_sample_pdf(dirty_pdf, 3, with_blanks=True)

        res = rotate_and_clean_pdf(dirty_pdf, cleaned_pdf, rotation=90, remove_blank=True)
        self.assertTrue(res["success"])
        self.assertEqual(res["initial_pages"], 3)
        self.assertEqual(res["final_pages"], 2)
        self.assertEqual(res["removed_blanks_count"], 1)
        self.assertEqual(res["rotation_applied"], 90)

        # Verify rotation of remaining pages in output
        doc = fitz.open(cleaned_pdf)
        self.assertEqual(len(doc), 2)
        self.assertEqual(doc[0].rotation, 90)
        self.assertEqual(doc[1].rotation, 90)
        doc.close()

    def test_stamp_watermark_pdf(self):
        """Test stamping diagonal watermark text across all pages."""
        src_pdf = os.path.join(self.scratch_dir, "pre_watermark.pdf")
        stamped_pdf = os.path.join(self.scratch_dir, "watermarked.pdf")

        self._create_sample_pdf(src_pdf, 2)
        out_path = stamp_watermark_pdf(src_pdf, stamped_pdf, watermark_text="TÀI LIỆU MẬT")
        self.assertTrue(os.path.isfile(out_path))

        doc = fitz.open(out_path)
        for page in doc:
            page_text = page.get_text().replace("\xa0", " ")
            self.assertIn("TÀI LIỆU MẬT", page_text)
        doc.close()

    def test_compress_pdf(self):
        """Test garbage collection and stream deflation for PDF compression."""
        src_pdf = os.path.join(self.scratch_dir, "pre_compress.pdf")
        comp_pdf = os.path.join(self.scratch_dir, "compressed.pdf")

        self._create_sample_pdf(src_pdf, 4)
        res = compress_pdf(src_pdf, comp_pdf)

        self.assertTrue(res["success"])
        self.assertTrue(os.path.isfile(res["output_path"]))
        self.assertGreater(res["original_bytes"], 0)
        self.assertGreater(res["compressed_bytes"], 0)


if __name__ == "__main__":
    unittest.main()
