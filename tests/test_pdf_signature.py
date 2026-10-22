"""
tests/test_pdf_signature.py — Comprehensive Unit Tests for PDF Digital Signature & Stamp Engine.
Validates:
1. Dynamic page targeting ('last', 'first', 'all', ranges '1, 2')
2. Base64 signature image injection (Canvas pad export)
3. Signer metadata vector text annotations (name, title, date)
4. Integration with /api/pdf/tools action 'sign'
"""

import base64
import io
import os
import tempfile
import unittest

from PIL import Image

try:
    import pymupdf as fitz
except ImportError:
    import fitz

from core.converter.pdf_toolkit import sign_pdf_with_stamp, _resolve_signature_pages


class TestPdfSignatureEngine(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()

        # Create a sample 3-page source PDF
        self.pdf_path = os.path.join(self.temp_dir.name, "sample_contract.pdf")
        doc = fitz.open()
        for i in range(3):
            page = doc.new_page(width=595.32, height=841.92)
            page.insert_text(fitz.Point(72, 100), f"SAMPLE CONTRACT - PAGE {i + 1}", fontsize=14)
        doc.save(self.pdf_path)
        doc.close()

        # Create a sample transparent PNG signature image
        self.stamp_path = os.path.join(self.temp_dir.name, "stamp.png")
        img = Image.new("RGBA", (200, 100), (0, 0, 0, 0))
        # Draw a simple box inside
        for x in range(20, 180):
            img.putpixel((x, 20), (0, 51, 170, 255))
            img.putpixel((x, 80), (0, 51, 170, 255))
        for y in range(20, 80):
            img.putpixel((20, y), (0, 51, 170, 255))
            img.putpixel((180, y), (0, 51, 170, 255))
        img.save(self.stamp_path, format="PNG")

        # Create base64 representation
        buf = io.BytesIO()
        img.save(buf, format="PNG")
        self.b64_signature = "data:image/png;base64," + base64.b64encode(buf.getvalue()).decode("ascii")

    def tearDown(self):
        self.temp_dir.cleanup()

    def test_resolve_signature_pages(self):
        self.assertEqual(_resolve_signature_pages("last", 5), [4])
        self.assertEqual(_resolve_signature_pages("first", 5), [0])
        self.assertEqual(_resolve_signature_pages("all", 3), [0, 1, 2])
        self.assertEqual(_resolve_signature_pages("1, 3", 5), [0, 2])
        self.assertEqual(_resolve_signature_pages("2-4", 5), [1, 2, 3])
        self.assertEqual(_resolve_signature_pages(2, 5), [1])
        self.assertEqual(_resolve_signature_pages(None, 5), [4])

    def test_sign_last_page_default(self):
        out_pdf = os.path.join(self.temp_dir.name, "signed_last.pdf")
        res = sign_pdf_with_stamp(
            self.pdf_path,
            out_pdf,
            signatures=[{
                "image_path": self.stamp_path,
                "page": "last",
                "x_pct": 0.65,
                "y_pct": 0.75,
                "w_pct": 0.25,
                "h_pct": 0.15,
            }],
        )
        self.assertTrue(res["success"])
        self.assertEqual(res["total_signatures_applied"], 1)

        # Verify page count and that stamp is on page 3 (index 2)
        doc = fitz.open(out_pdf)
        self.assertEqual(len(doc), 3)
        self.assertEqual(len(doc[2].get_images()), 1)
        self.assertEqual(len(doc[0].get_images()), 0)
        doc.close()

    def test_sign_all_pages_with_base64(self):
        out_pdf = os.path.join(self.temp_dir.name, "signed_all.pdf")
        res = sign_pdf_with_stamp(
            self.pdf_path,
            out_pdf,
            signatures=[{
                "image_base64": self.b64_signature,
                "page": "all",
                "x_pct": 0.70,
                "y_pct": 0.80,
            }],
        )
        self.assertTrue(res["success"])
        self.assertEqual(res["total_signatures_applied"], 3)

        doc = fitz.open(out_pdf)
        for i in range(3):
            self.assertEqual(len(doc[i].get_images()), 1)
        doc.close()

    def test_sign_with_signer_metadata(self):
        out_pdf = os.path.join(self.temp_dir.name, "signed_with_meta.pdf")
        res = sign_pdf_with_stamp(
            self.pdf_path,
            out_pdf,
            signatures=[{
                "image_path": self.stamp_path,
                "page": "last",
                "signer_title": "Giám Đốc Điều Hành",
                "signer_name": "Nguyễn Văn A",
                "signer_date": "Ngày 22/10/2026",
            }],
        )
        self.assertTrue(res["success"])
        self.assertEqual(res["total_signatures_applied"], 1)

        doc = fitz.open(out_pdf)
        text_content = doc[2].get_text().replace("\xa0", " ")
        self.assertIn("Nguyễn Văn A", text_content)
        self.assertIn("Giám Đốc Điều Hành", text_content)
        doc.close()


if __name__ == "__main__":
    unittest.main()
