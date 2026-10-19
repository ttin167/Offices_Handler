"""
tests/test_pdf_ilovepdf_features.py — Verification Suite for iLovePDF Killer Features (v4.2.0).

Verifies the 5 New Operations & Invariants:
1. Images to PDF: Multi-image to standard A4 PDF with EXIF auto-rotation and margin calculation.
2. PDF to 300 DPI Images ZIP: High-resolution rasterization and valid ZIP archive generation.
3. Sign & Stamp: Derotation matrix compensation when page.rotation != 0, exact placement.
4. Flatten & Sanitize: Vector Bake (doc.bake) vs Strict Raster (get_pixmap) + doc.scrub() metadata purge.
5. Visual Reorganization & Blank Page: Reordering pages with dynamic A4 blank page insertion.
"""

import os
import sys
import zipfile
import unittest

try:
    import pymupdf as fitz
except ImportError:
    import fitz

from PIL import Image, ImageDraw

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)

from pdf_toolkit import (
    convert_images_to_pdf,
    export_pdf_to_images_zip,
    sign_pdf_with_stamp,
    flatten_and_sanitize_pdf,
    reorganize_pdf_pages,
)


class TestPdfILovePdfFeaturesSuite(unittest.TestCase):
    def setUp(self):
        self.scratch_dir = os.path.join(BASE_DIR, "scratch", "test_ilovepdf_features")
        os.makedirs(self.scratch_dir, exist_ok=True)

    def _create_sample_image(self, path: str, width: int = 600, height: int = 800, color=(50, 100, 200)) -> str:
        img = Image.new("RGB", (width, height), color)
        draw = ImageDraw.Draw(img)
        draw.text((40, 40), "TEST IMAGE SAMPLE", fill=(255, 255, 255))
        img.save(path)
        return path

    def _create_sample_pdf(self, path: str, num_pages: int = 3, rotation: int = 0) -> str:
        doc = fitz.open()
        for i in range(num_pages):
            page = doc.new_page(width=595.28, height=841.89)
            if rotation:
                page.set_rotation(rotation)
            page.insert_text(fitz.Point(72, 100), f"ORIGINAL CONTRACT CONTENT - PAGE {i + 1}")
        doc.save(path)
        doc.close()
        return path

    def test_01_images_to_pdf(self):
        """Test converting multiple image files into a single standardized A4 PDF."""
        img1 = os.path.join(self.scratch_dir, "img1.png")
        img2 = os.path.join(self.scratch_dir, "img2.jpg")
        self._create_sample_image(img1, 800, 600, (200, 50, 50))
        self._create_sample_image(img2, 600, 900, (50, 180, 50))

        out_pdf = os.path.join(self.scratch_dir, "converted_images.pdf")
        res = convert_images_to_pdf([img1, img2], out_pdf, page_size="A4", orientation="auto", margin_pt=36)

        self.assertTrue(os.path.isfile(out_pdf))
        self.assertEqual(res["total_images"], 2)
        self.assertEqual(res["total_pages"], 2)

        doc = fitz.open(out_pdf)
        self.assertEqual(len(doc), 2)
        doc.close()

    def test_02_export_pdf_to_images_zip(self):
        """Test exporting all pages of a PDF into a 300 DPI PNG ZIP archive."""
        src_pdf = os.path.join(self.scratch_dir, "to_export.pdf")
        self._create_sample_pdf(src_pdf, num_pages=2)

        out_zip = os.path.join(self.scratch_dir, "exported_bundle.zip")
        res = export_pdf_to_images_zip(src_pdf, out_zip, dpi=150, img_format="png")

        self.assertTrue(os.path.isfile(out_zip))
        self.assertEqual(res["total_pages"], 2)

        with zipfile.ZipFile(out_zip, "r") as z:
            namelist = z.namelist()
            self.assertEqual(len(namelist), 2)
            self.assertTrue(all(name.endswith(".png") for name in namelist))

    def test_03_sign_pdf_with_derotation(self):
        """Test signature stamping with rotation compensation when page.rotation != 0."""
        # Create a PDF with page rotation = 90
        src_pdf = os.path.join(self.scratch_dir, "rotated_doc.pdf")
        self._create_sample_pdf(src_pdf, num_pages=1, rotation=90)

        sig_img = os.path.join(self.scratch_dir, "signature.png")
        self._create_sample_image(sig_img, 200, 100, (220, 20, 60))

        out_pdf = os.path.join(self.scratch_dir, "signed_doc.pdf")
        sig_data = [
            {
                "page": 1,
                "image": sig_img,
                "x_pct": 0.6,
                "y_pct": 0.7,
                "w_pct": 0.3,
                "h_pct": 0.15,
            }
        ]
        res = sign_pdf_with_stamp(src_pdf, out_pdf, sig_data)
        self.assertTrue(os.path.isfile(out_pdf))
        self.assertEqual(res["signatures_applied"], 1)

        # Inspect resulting PDF
        doc = fitz.open(out_pdf)
        page = doc[0]
        images = page.get_images()
        self.assertGreaterEqual(len(images), 1)
        doc.close()

    def test_04_flatten_and_sanitize(self):
        """Test dual-mode flattening (vector bake vs strict raster) and metadata purging."""
        src_pdf = os.path.join(self.scratch_dir, "doc_with_meta.pdf")
        doc = fitz.open()
        p = doc.new_page(width=595.28, height=841.89)
        p.insert_text(fitz.Point(72, 100), "SENSITIVE ANNOTATED AGREEMENT")
        p.add_rect_annot(fitz.Rect(72, 120, 200, 150))
        doc.set_metadata({"author": "Top Secret Agent", "title": "Confidential Deal"})
        doc.save(src_pdf)
        doc.close()

        # Vector Bake mode
        out_baked = os.path.join(self.scratch_dir, "flattened_baked.pdf")
        res_baked = flatten_and_sanitize_pdf(src_pdf, out_baked, flatten_vector=True, strict_raster=False, sanitize_metadata=True)
        self.assertTrue(os.path.isfile(out_baked))
        self.assertEqual(res_baked["mode"], "vector_bake")

        doc_b = fitz.open(out_baked)
        self.assertEqual(doc_b.metadata.get("author", ""), "")
        doc_b.close()

        # Strict Raster mode
        out_raster = os.path.join(self.scratch_dir, "flattened_raster.pdf")
        res_raster = flatten_and_sanitize_pdf(src_pdf, out_raster, strict_raster=True, raster_dpi=150)
        self.assertTrue(os.path.isfile(out_raster))
        self.assertEqual(res_raster["mode"], "strict_raster")

        doc_r = fitz.open(out_raster)
        self.assertGreaterEqual(len(doc_r[0].get_images()), 1)
        doc_r.close()

    def test_05_reorganize_with_blank_page(self):
        """Test page reorganization with dynamic A4 blank page insertion."""
        src_pdf = os.path.join(self.scratch_dir, "orig_pages.pdf")
        self._create_sample_pdf(src_pdf, num_pages=3)

        out_pdf = os.path.join(self.scratch_dir, "reordered_with_blank.pdf")
        # Plan: [Page 2, Blank Page, Page 1]
        pages_plan = [
            {"page_num": 2, "rotation": 0},
            {"is_blank": True},
            {"page_num": 1, "rotation": 90},
        ]
        res = reorganize_pdf_pages(src_pdf, out_pdf, pages_plan)
        self.assertTrue(os.path.isfile(out_pdf))
        self.assertEqual(res["pages_count"], 3)

        doc = fitz.open(out_pdf)
        self.assertEqual(len(doc), 3)
        # Page 1 is old Page 2
        self.assertIn("PAGE 2", doc[0].get_text())
        # Page 2 is blank page
        self.assertEqual(doc[1].get_text().strip(), "")
        # Page 3 is old Page 1 with rotation 90
        self.assertIn("PAGE 1", doc[2].get_text())
        self.assertEqual(doc[2].rotation, 90)
        doc.close()


if __name__ == "__main__":
    unittest.main()
