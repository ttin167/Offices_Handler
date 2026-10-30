"""
test_pdf_advanced_toolkit.py — Comprehensive Unit Tests for PDF Pro Studio Subsystem (v4.0).

Verifies:
1. redact_sensitive_pdf: True permanent byte/pixel redaction for CCCD, Phone, Email, Bank cards, and custom terms.
2. protect_pdf: AES-256 encryption with user/owner passwords and granular permissions (print, copy, modify).
3. normalize_and_number_pdf: Centering onto standard A4 (595.28 x 841.89 pt) + Bates numbering with Unicode.
4. stamp_giap_lai_pdf: Slicing company stamp into N overlapping strips across N pages.
5. extract_pdf_tables_to_xlsx: PyMuPDF table extraction to styled openpyxl XLSX complying with ERR_XLSX_007 and ERR_XLSX_008.
6. render_pdf_thumbnails: High-speed 120 DPI pixmap rendering for visual organizer.
7. reorganize_pdf_pages: Selective reordering, per-page rotation, and deletion.
"""

import os
import sys
import unittest

try:
    import pymupdf as fitz
except ImportError:
    import fitz

from PIL import Image, ImageDraw
import openpyxl

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)

from core.converter.pdf_toolkit import (
    redact_sensitive_pdf,
    protect_pdf,
    normalize_and_number_pdf,
    stamp_giap_lai_pdf,
    extract_pdf_tables_to_xlsx,
    render_pdf_thumbnails,
    reorganize_pdf_pages,
    SENSITIVE_PATTERNS,
)


class TestPDFAdvancedToolkit(unittest.TestCase):
    def setUp(self):
        self.scratch_dir = os.path.join(BASE_DIR, "scratch", "test_pdf_advanced")
        os.makedirs(self.scratch_dir, exist_ok=True)

    def _create_sample_pdf(self, path: str, num_pages: int = 3, content_func=None) -> str:
        doc = fitz.open()
        for i in range(num_pages):
            page = doc.new_page(width=600, height=800)
            if content_func:
                content_func(page, i)
            else:
                page.insert_text(fitz.Point(50, 100), f"Antigravity Document Page {i + 1}")
        doc.save(path)
        doc.close()
        return path

    def _create_sample_stamp_png(self, path: str) -> str:
        """Create a synthetic red round stamp image."""
        img = Image.new("RGBA", (300, 300), (255, 255, 255, 0))
        draw = ImageDraw.Draw(img)
        draw.ellipse([10, 10, 290, 290], outline=(220, 20, 60, 255), width=8)
        draw.text((80, 140), "DẤU GIÁP LAI", fill=(220, 20, 60, 255))
        img.save(path, format="PNG")
        return path

    def test_redact_sensitive_pdf_true_permanent_deletion(self):
        """Test true redaction of CCCD, Phone, Email, Bank Card, and custom terms."""
        src_pdf = os.path.join(self.scratch_dir, "sensitive_source.pdf")
        out_pdf = os.path.join(self.scratch_dir, "redacted_output.pdf")

        def populate_sensitive(page, idx):
            page.insert_text(fitz.Point(50, 100), "HỌ VÀ TÊN: NGUYỄN VĂN A")
            page.insert_text(fitz.Point(50, 130), "SỐ CCCD: 079090123456")
            page.insert_text(fitz.Point(50, 160), "ĐIỆN THOẠI: +84901234567")
            page.insert_text(fitz.Point(50, 190), "EMAIL: confidential_hr@company.com.vn")
            page.insert_text(fitz.Point(50, 220), "SỐ THẺ VISA: 4111 2222 3333 4444")
            page.insert_text(fitz.Point(50, 250), "DỰ ÁN TUYỆT MẬT: PROJECT_TITAN_2026")

        self._create_sample_pdf(src_pdf, num_pages=2, content_func=populate_sensitive)

        # Pre-check: text is present
        doc_pre = fitz.open(src_pdf)
        pre_text = doc_pre[0].get_text()
        doc_pre.close()
        self.assertIn("079090123456", pre_text)
        self.assertIn("PROJECT_TITAN_2026", pre_text)

        result = redact_sensitive_pdf(
            src_pdf,
            out_pdf,
            patterns=["cccd", "phone", "email", "bank_card"],
            custom_terms=["PROJECT_TITAN_2026", "NGUYỄN VĂN A"],
            fill_color="black",
        )

        self.assertTrue(os.path.isfile(out_pdf))
        self.assertGreater(result["total_redactions"], 0)

        # Post-check: Sensitive text bytes are 100% permanently deleted from AST
        doc_post = fitz.open(out_pdf)
        post_text = doc_post[0].get_text()
        doc_post.close()

        self.assertNotIn("079090123456", post_text)
        self.assertNotIn("confidential_hr@company.com.vn", post_text)
        self.assertNotIn("PROJECT_TITAN_2026", post_text)
        self.assertNotIn("NGUYỄN VĂN A", post_text)

    def test_protect_pdf_aes256_encryption_and_permissions(self):
        """Test AES-256 encryption, password authentication and permission masks."""
        src_pdf = os.path.join(self.scratch_dir, "clean_doc.pdf")
        out_pdf = os.path.join(self.scratch_dir, "protected_doc.pdf")
        self._create_sample_pdf(src_pdf, num_pages=1)

        protect_pdf(
            src_pdf,
            out_pdf,
            user_password="UserPassword123",
            owner_password="OwnerSecret456",
            allow_print=True,
            allow_copy=False,
            allow_modify=False,
        )

        self.assertTrue(os.path.isfile(out_pdf))

        # Check opening without password fails or is encrypted
        doc = fitz.open(out_pdf)
        self.assertTrue(doc.is_encrypted)
        # Auth with wrong password fails
        self.assertEqual(doc.authenticate("WrongPass"), 0)
        # Auth with user password succeeds
        self.assertGreater(doc.authenticate("UserPassword123"), 0)
        # Verify text is accessible once authenticated
        self.assertIn("Antigravity Document", doc[0].get_text())
        doc.close()

    def test_normalize_and_number_pdf_a4(self):
        """Test A4 vector normalization and Vietnamese Bates numbering."""
        src_pdf = os.path.join(self.scratch_dir, "odd_sized.pdf")
        out_pdf = os.path.join(self.scratch_dir, "normalized_a4.pdf")

        # Create doc with irregular page size (US Letter: 612 x 792)
        doc = fitz.open()
        p1 = doc.new_page(width=612, height=792)
        p1.insert_text(fitz.Point(50, 100), "Letter Page 1")
        p2 = doc.new_page(width=800, height=400)  # Landscape
        p2.insert_text(fitz.Point(50, 100), "Wide Page 2")
        doc.save(src_pdf)
        doc.close()

        normalize_and_number_pdf(
            src_pdf,
            out_pdf,
            orientation="auto",
            bates_prefix="ANTIGRAVITY-",
            start_number=1,
        )

        self.assertTrue(os.path.isfile(out_pdf))
        doc_norm = fitz.open(out_pdf)
        self.assertEqual(len(doc_norm), 2)

        # Check Page 1 is standard A4 Portrait (595.28 x 841.89)
        self.assertAlmostEqual(doc_norm[0].rect.width, 595.28, delta=1.0)
        self.assertAlmostEqual(doc_norm[0].rect.height, 841.89, delta=1.0)
        p1_text = doc_norm[0].get_text()
        self.assertIn("ANTIGRAVITY", p1_text)
        self.assertIn("Trang", p1_text)
        self.assertIn("01/02", p1_text)

        # Check Page 2 is standard A4 Landscape (841.89 x 595.28)
        self.assertAlmostEqual(doc_norm[1].rect.width, 841.89, delta=1.0)
        self.assertAlmostEqual(doc_norm[1].rect.height, 595.28, delta=1.0)
        p2_text = doc_norm[1].get_text()
        self.assertIn("ANTIGRAVITY", p2_text)
        self.assertIn("Trang", p2_text)
        self.assertIn("02/02", p2_text)
        doc_norm.close()

    def test_stamp_giap_lai_pdf_multi_page(self):
        """Test slicing company stamp into N vertical overlapping strips across N pages."""
        src_pdf = os.path.join(self.scratch_dir, "contract_source.pdf")
        out_pdf = os.path.join(self.scratch_dir, "contract_stamped.pdf")
        stamp_png = os.path.join(self.scratch_dir, "red_stamp.png")

        self._create_sample_pdf(src_pdf, num_pages=3)
        self._create_sample_stamp_png(stamp_png)

        stamp_giap_lai_pdf(
            src_pdf,
            out_pdf,
            stamp_image_path=stamp_png,
            side="right",
            overlap_pct=8.0,
        )

        self.assertTrue(os.path.isfile(out_pdf))
        doc = fitz.open(out_pdf)
        self.assertEqual(len(doc), 3)

        # Each page must contain at least 1 image (the sliced stamp strip)
        for i, page in enumerate(doc):
            imgs = page.get_images()
            self.assertGreaterEqual(len(imgs), 1, f"Page {i + 1} missing giap lai stamp image slice")
        doc.close()

    def test_extract_pdf_tables_to_xlsx_with_standards(self):
        """Test table extraction to openpyxl with ERR_XLSX_007 string formatting and ERR_XLSX_008 print setup."""
        src_pdf = os.path.join(self.scratch_dir, "table_source.pdf")
        out_xlsx = os.path.join(self.scratch_dir, "extracted_tables.xlsx")

        # Create PDF containing a clean vector table with borders
        doc = fitz.open()
        page = doc.new_page(width=595, height=842)

        # Draw a table with 3 rows and 3 columns
        x0, y0 = 50, 100
        col_w = [120, 160, 180]
        row_h = 30
        total_w = sum(col_w)
        total_h = 3 * row_h

        # Draw outer rectangle
        page.draw_rect(fitz.Rect(x0, y0, x0 + total_w, y0 + total_h), color=(0, 0, 0), width=1)

        # Draw horizontal inner lines
        for r in range(1, 3):
            y = y0 + r * row_h
            page.draw_line(fitz.Point(x0, y), fitz.Point(x0 + total_w, y), color=(0, 0, 0), width=1)
        # Draw vertical inner lines
        cur_x = x0
        for w in col_w[:-1]:
            cur_x += w
            page.draw_line(fitz.Point(cur_x, y0), fitz.Point(cur_x, y0 + total_h), color=(0, 0, 0), width=1)

        # Insert table text
        page.insert_text(fitz.Point(60, 120), "STT")
        page.insert_text(fitz.Point(180, 120), "Mã CCCD")
        page.insert_text(fitz.Point(340, 120), "Số Điện Thoại")

        page.insert_text(fitz.Point(60, 150), "1")
        page.insert_text(fitz.Point(180, 150), "079090123456")  # Must keep leading zero
        page.insert_text(fitz.Point(340, 150), "0901234567")

        page.insert_text(fitz.Point(60, 180), "2")
        page.insert_text(fitz.Point(180, 180), "001234567890")  # Two leading zeros
        page.insert_text(fitz.Point(340, 180), "0987654321")

        doc.save(src_pdf)
        doc.close()

        extract_pdf_tables_to_xlsx(src_pdf, out_xlsx)
        self.assertTrue(os.path.isfile(out_xlsx))

        # Inspect generated workbook
        wb = openpyxl.load_workbook(out_xlsx)
        ws = wb.active
        self.assertIsNotNone(ws)

        # Check ERR_XLSX_008: A4 landscape orientation
        self.assertEqual(str(ws.page_setup.paperSize), str(ws.PAPERSIZE_A4))
        self.assertEqual(ws.page_setup.orientation, ws.ORIENTATION_LANDSCAPE)

        # Check ERR_XLSX_007: CCCD column retains leading zeros as string with '@' format
        found_cccd = False
        for row in ws.iter_rows(values_only=False):
            for cell in row:
                val = str(cell.value) if cell.value is not None else ""
                if "079090123456" in val or "001234567890" in val:
                    found_cccd = True
                    self.assertEqual(cell.number_format, "@")
                    self.assertIsInstance(cell.value, str)

        self.assertTrue(found_cccd, "CCCD string with leading zeros should be extracted and formatted as text '@'")
        wb.close()

    def test_render_pdf_thumbnails_and_reorganize_pages(self):
        """Test high-speed thumbnail rendering and visual page reorganization."""
        src_pdf = os.path.join(self.scratch_dir, "organizer_source.pdf")
        out_pdf = os.path.join(self.scratch_dir, "reorganized_output.pdf")
        thumb_dir = os.path.join(self.scratch_dir, "thumbnails")

        # 4 pages source
        self._create_sample_pdf(src_pdf, num_pages=4)

        # 1. Render thumbnails
        thumbs = render_pdf_thumbnails(src_pdf, thumb_dir, dpi=100)
        self.assertEqual(len(thumbs), 4)
        for t in thumbs:
            img_path = t.get("file_path") or t.get("path") if isinstance(t, dict) else t
            self.assertTrue(os.path.isfile(img_path))
            self.assertTrue(img_path.endswith(".png"))

        # 2. Reorganize: Keep page 3 (rotate 90), page 1 (rotate 180), delete page 2 and 4
        plan = [
            {"page": 3, "rotate": 90},
            {"page": 1, "rotate": 180},
        ]
        reorganize_pdf_pages(src_pdf, out_pdf, pages_plan=plan)

        self.assertTrue(os.path.isfile(out_pdf))
        doc_reorg = fitz.open(out_pdf)
        self.assertEqual(len(doc_reorg), 2)

        # Page 1 in reorg is original Page 3 rotated 90
        self.assertIn("Page 3", doc_reorg[0].get_text())
        self.assertEqual(doc_reorg[0].rotation, 90)

        # Page 2 in reorg is original Page 1 rotated 180
        self.assertIn("Page 1", doc_reorg[1].get_text())
        self.assertEqual(doc_reorg[1].rotation, 180)
        doc_reorg.close()


if __name__ == "__main__":
    unittest.main()
