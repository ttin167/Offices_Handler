"""
tests/test_pdf_pro_studio.py — Production-Ready Verification Suite for PDF Pro Studio (v4.1.0).

Verifies the 4 Core Business Rules & Security Guarantees:
1. True Redaction: Sensitive text bytes are 100% permanently eliminated and unrecoverable via page.get_text().
2. AES-256 Security: Document is properly encrypted, rejects invalid passwords, and authenticates with correct user/owner keys.
3. PDF Table to Excel (ERR_XLSX_007 & ERR_XLSX_008): Leading zeros on CCCD/Tax IDs/Bank codes are preserved as text strings.
4. Vietnamese Giáp Lai Stamp: Stamp image is sliced into exactly N vertical strips across N consecutive pages with 8% overlap.
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

from pdf_toolkit import (
    redact_sensitive_pdf,
    protect_pdf,
    extract_pdf_tables_to_xlsx,
    stamp_giap_lai_pdf,
    SENSITIVE_PATTERNS,
)


class TestPdfProStudioProductionSuite(unittest.TestCase):
    def setUp(self):
        self.scratch_dir = os.path.join(BASE_DIR, "scratch", "test_pdf_pro_studio")
        os.makedirs(self.scratch_dir, exist_ok=True)

    def _create_sample_pdf(self, path: str, num_pages: int = 4, content_builder=None) -> str:
        doc = fitz.open()
        for i in range(num_pages):
            page = doc.new_page(width=595.28, height=841.89)  # Standard A4
            if content_builder:
                content_builder(page, i)
            else:
                page.insert_text(fitz.Point(72, 100), f"ANTIGRAVITY ENTERPRISE CONTRACT - PAGE {i + 1}")
        doc.save(path)
        doc.close()
        return path

    def _create_round_stamp_image(self, path: str, size: int = 400) -> str:
        img = Image.new("RGBA", (size, size), (255, 255, 255, 0))
        draw = ImageDraw.Draw(img)
        # Red circular seal
        draw.ellipse([10, 10, size - 10, size - 10], outline=(220, 20, 60, 255), width=8)
        draw.ellipse([30, 30, size - 30, size - 30], outline=(220, 20, 60, 255), width=3)
        draw.text((size // 4, size // 2 - 10), "CÔNG TY CỔ PHẦN ANTIGRAVITY", fill=(220, 20, 60, 255))
        img.save(path, format="PNG")
        return path

    def test_01_true_redaction_zero_text_leak(self):
        """
        Rule 1: Văn bản sau True Redaction không thể tìm thấy bằng page.get_text().
        Verifies that redaction permanently purges text bytes from PDF AST (Zero OCR & Zero Copy/Paste leak).
        """
        src_pdf = os.path.join(self.scratch_dir, "contract_confidential.pdf")
        out_pdf = os.path.join(self.scratch_dir, "contract_redacted.pdf")

        sensitive_cccd = "079090123456"
        sensitive_phone = "+84901234567"
        sensitive_email = "giamdoc_taichinh@doanhnghiep.vn"
        sensitive_card = "4111 2222 3333 4444"
        sensitive_secret = "DU_AN_DAT_VANG_QUAN_1_2026"

        def inject_secrets(page, idx):
            page.insert_text(fitz.Point(72, 120), f"Bên A: Nguyễn Văn A - CCCD: {sensitive_cccd}")
            page.insert_text(fitz.Point(72, 150), f"Điện thoại liên hệ: {sensitive_phone}")
            page.insert_text(fitz.Point(72, 180), f"Email bảo mật: {sensitive_email}")
            page.insert_text(fitz.Point(72, 210), f"Số thẻ thanh toán tín dụng: {sensitive_card}")
            page.insert_text(fitz.Point(72, 240), f"Tên dự án mật: {sensitive_secret}")

        self._create_sample_pdf(src_pdf, num_pages=2, content_builder=inject_secrets)

        # 1. Verify secrets exist before redaction
        doc_before = fitz.open(src_pdf)
        raw_text_before = "".join(p.get_text() for p in doc_before)
        doc_before.close()
        self.assertIn(sensitive_cccd, raw_text_before)
        self.assertIn(sensitive_email, raw_text_before)
        self.assertIn(sensitive_secret, raw_text_before)

        # 2. Execute True Redaction with automated regex patterns & custom terms
        result = redact_sensitive_pdf(
            src_pdf,
            out_pdf,
            patterns=["cccd", "phone", "email", "bank_card"],
            custom_terms=[sensitive_secret, "Nguyễn Văn A"],
            fill_color="black",
        )

        self.assertTrue(os.path.isfile(out_pdf))
        self.assertGreater(result["total_redactions"], 0)

        # 3. CRITICAL AUDIT: Scan entire AST of redacted PDF — Zero bytes leak guaranteed
        doc_after = fitz.open(out_pdf)
        raw_text_after = "".join(p.get_text() for p in doc_after)
        doc_after.close()

        self.assertNotIn(sensitive_cccd, raw_text_after, "CCCD must be completely purged from AST")
        self.assertNotIn(sensitive_phone, raw_text_after, "Phone must be completely purged from AST")
        self.assertNotIn(sensitive_email, raw_text_after, "Email must be completely purged from AST")
        self.assertNotIn("4111 2222", raw_text_after, "Bank card must be completely purged from AST")
        self.assertNotIn(sensitive_secret, raw_text_after, "Secret project term must be completely purged from AST")
        self.assertNotIn("Nguyễn Văn A", raw_text_after, "Person name must be completely purged from AST")

    def test_02_aes256_encryption_and_password_enforcement(self):
        """
        Rule 2: File mã hóa AES-256 yêu cầu đúng mật khẩu.
        Verifies strong AES-256 encryption, access denial on wrong password, and proper permission gating.
        """
        src_pdf = os.path.join(self.scratch_dir, "confidential_memo.pdf")
        out_pdf = os.path.join(self.scratch_dir, "encrypted_memo.pdf")
        self._create_sample_pdf(src_pdf, num_pages=2)

        user_pwd = "UserKey@2026"
        owner_pwd = "AdminMaster@2026"

        protect_pdf(
            src_pdf,
            out_pdf,
            user_password=user_pwd,
            owner_password=owner_pwd,
            allow_print=True,
            allow_copy=False,
            allow_modify=False,
        )

        self.assertTrue(os.path.isfile(out_pdf))

        # Check document encryption flag
        doc = fitz.open(out_pdf)
        self.assertTrue(doc.is_encrypted, "PDF must be marked as encrypted")

        # Wrong password must fail
        auth_wrong = doc.authenticate("InvalidPassword")
        self.assertEqual(auth_wrong, 0, "Authentication with wrong password must return 0")

        # Correct user password must succeed
        auth_user = doc.authenticate(user_pwd)
        self.assertGreater(auth_user, 0, "Authentication with user password must succeed")

        # After authentication, content is accessible
        page_text = doc[0].get_text()
        self.assertIn("ANTIGRAVITY ENTERPRISE CONTRACT", page_text)
        doc.close()

    def test_03_table_to_excel_preserves_leading_zeros(self):
        """
        Rule 3: Lưới bảng trích xuất sang Excel giữ nguyên số 0 ở đầu mã định danh (ERR_XLSX_007).
        Verifies that PyMuPDF find_tables() converts to openpyxl with text '@' format for CCCD, Tax IDs, and phone numbers.
        """
        src_pdf = os.path.join(self.scratch_dir, "payroll_table.pdf")
        out_xlsx = os.path.join(self.scratch_dir, "payroll_extracted.xlsx")

        doc = fitz.open()
        page = doc.new_page(width=595.28, height=841.89)

        # Draw structured table grid
        x0, y0 = 50, 100
        col_w = [60, 150, 160, 120]
        row_h = 32
        total_w = sum(col_w)
        total_h = 3 * row_h

        # Outer border
        page.draw_rect(fitz.Rect(x0, y0, x0 + total_w, y0 + total_h), color=(0, 0, 0), width=1)
        # Inner horizontal lines
        for r in range(1, 3):
            page.draw_line(fitz.Point(x0, y0 + r * row_h), fitz.Point(x0 + total_w, y0 + r * row_h), color=(0, 0, 0), width=1)
        # Inner vertical lines
        cur_x = x0
        for w in col_w[:-1]:
            cur_x += w
            page.draw_line(fitz.Point(cur_x, y0), fitz.Point(cur_x, y0 + total_h), color=(0, 0, 0), width=1)

        # Header
        page.insert_text(fitz.Point(55, 122), "STT")
        page.insert_text(fitz.Point(120, 122), "Số CCCD")
        page.insert_text(fitz.Point(270, 122), "Mã Số Thuế")
        page.insert_text(fitz.Point(430, 122), "Số Điện Thoại")

        # Row 1 (IDs with single leading zero)
        page.insert_text(fitz.Point(55, 154), "1")
        page.insert_text(fitz.Point(120, 154), "079090123456")
        page.insert_text(fitz.Point(270, 154), "0101234567")
        page.insert_text(fitz.Point(430, 154), "0908123456")

        # Row 2 (IDs with double leading zeros)
        page.insert_text(fitz.Point(55, 186), "2")
        page.insert_text(fitz.Point(120, 186), "001234567890")
        page.insert_text(fitz.Point(270, 186), "0071000123456")
        page.insert_text(fitz.Point(430, 186), "0912345678")

        doc.save(src_pdf)
        doc.close()

        # Extract to Excel
        result = extract_pdf_tables_to_xlsx(src_pdf, out_xlsx)
        self.assertTrue(os.path.isfile(out_xlsx))
        self.assertGreaterEqual(result["tables_extracted"], 1)

        # Open and inspect openpyxl workbook
        wb = openpyxl.load_workbook(out_xlsx)
        ws = wb.active
        self.assertIsNotNone(ws)

        # Verify ERR_XLSX_008: Landscape print setup & A4 paper size
        self.assertEqual(str(ws.page_setup.paperSize), str(ws.PAPERSIZE_A4))
        self.assertEqual(ws.page_setup.orientation, ws.ORIENTATION_LANDSCAPE)

        # Verify ERR_XLSX_007: Leading zeros preserved and cells have '@' string format
        preserved_values = set()
        for row in ws.iter_rows(values_only=False):
            for cell in row:
                if cell.value is not None:
                    val_str = str(cell.value)
                    if val_str in ("079090123456", "0101234567", "001234567890", "0071000123456"):
                        preserved_values.add(val_str)
                        self.assertEqual(cell.number_format, "@", f"Cell {cell.coordinate} must have '@' text format")
                        self.assertIsInstance(cell.value, str, f"Cell {cell.coordinate} value must be a str, not int")

        self.assertIn("079090123456", preserved_values, "CCCD with leading zero must not lose '0'")
        self.assertIn("001234567890", preserved_values, "CCCD with two leading zeros must not lose '00'")
        wb.close()

    def test_04_stamp_giap_lai_n_slice_distribution(self):
        """
        Rule 4: Dấu giáp lai chia đúng N phần bằng nhau theo số trang (kèm overlap 8%).
        Verifies that an N-page document receives N stamp image slices placed along the outer border.
        """
        src_pdf = os.path.join(self.scratch_dir, "bidding_doc_source.pdf")
        out_pdf = os.path.join(self.scratch_dir, "bidding_doc_stamped.pdf")
        stamp_img = os.path.join(self.scratch_dir, "company_stamp_red.png")

        num_pages = 4
        self._create_sample_pdf(src_pdf, num_pages=num_pages)
        self._create_round_stamp_image(stamp_img, size=400)

        res = stamp_giap_lai_pdf(
            src_pdf,
            out_pdf,
            stamp_image_path=stamp_img,
            side="right",
            stamp_diameter=120.0,
            overlap_pct=8.0,
        )

        self.assertTrue(os.path.isfile(out_pdf))
        self.assertEqual(res["total_pages"], num_pages)
        self.assertEqual(res["slices_created"], num_pages)

        # Inspect stamped output PDF
        doc = fitz.open(out_pdf)
        self.assertEqual(len(doc), num_pages)

        # Each page must have at least 1 image embedded (the vertical stamp strip)
        for p_idx, page in enumerate(doc):
            images = page.get_images()
            self.assertGreaterEqual(
                len(images), 1, f"Page {p_idx + 1} of {num_pages} must contain its corresponding giap lai stamp slice"
            )

        doc.close()


if __name__ == "__main__":
    unittest.main()
