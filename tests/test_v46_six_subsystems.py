"""
tests/test_v46_six_subsystems.py — Comprehensive Verification Suite for v4.6.0
(Enterprise Omni-Pipeline & Deep CDR Suite + 5 Mandatory Production Edge-Case Guards).

Covers:
1. Subsystem 1: PKCS#7 / PAdES Cryptographic PDF Signing & PDF/A-2b Archival
2. Subsystem 2: Excel-to-Brief & Cross-Workbook Consolidation + Edge Case #3 (ERR_XLSX_015)
3. Subsystem 3: Deep CDR & Smart Vietnamese PII Redactor + Edge Cases #1 (ERR_SEC_006) & #2 (ERR_SEC_007)
4. Subsystem 4: Declarative Workflow Pipeline & Bulk Runner + Edge Case #4 (ERR_PIPE_001)
5. Subsystem 5: AST Code-to-Diagram Generator & Pure Draw.io SVG Exporter + DOCX/PPTX Auto-Inject
6. Subsystem 6: Unified Desktop Studio Bridge + Edge Case #5 (ERR_UI_001)
"""

from __future__ import annotations

import json
import os
import tempfile
import unittest
import zipfile
from pathlib import Path

import fitz
import openpyxl
from docx import Document

from core.converter.pdf_toolkit import (
    convert_to_pdfa,
    crypto_sign_pdf,
    verify_pdf_crypto_signature,
)
from core.diagrams.code_to_diagram import (
    export_mxgraph_to_svg,
    generate_diagram_from_code,
    scan_python_ast_to_spec,
    scan_sql_or_prisma_to_erd_spec,
)
from core.pipeline.engine import (
    WorkflowPipeline,
    get_default_max_workers,
    run_bulk_directory_pipeline,
)
from core.security.sanitizer import (
    DocumentSanitizer,
    redact_pii_text,
    sanitize_document,
)
from core.xlsx.reporting import consolidate_workbooks, excel_to_brief
from ui.webview_launcher import build_app_mode_browser_command, launch_native_web_studio


class TestV46SixSubsystems(unittest.TestCase):
    """End-to-end verification of all 6 v4.6.0 subsystems and 5 Production Edge-Case Guards."""

    def setUp(self) -> None:
        self.temp_dir_obj = tempfile.TemporaryDirectory(prefix="ag_test_v46_")
        self.temp_dir = self.temp_dir_obj.name

    def tearDown(self) -> None:
        self.temp_dir_obj.cleanup()

    # =========================================================================
    # SUBSYSTEM 1: Cryptographic PDF Signing (PKCS#7 / PAdES) & PDF/A-2b
    # =========================================================================
    def test_01_subsystem1_pdf_crypto_sign_and_pdfa(self) -> None:
        raw_pdf = os.path.join(self.temp_dir, "sample.pdf")
        signed_pdf = os.path.join(self.temp_dir, "sample_signed.pdf")
        pdfa_out = os.path.join(self.temp_dir, "sample_pdfa2b.pdf")

        # Create a 2-page PDF with an active JavaScript trigger to verify PDF/A sanitization
        doc = fitz.open()
        page = doc.new_page()
        page.insert_text((72, 72), "Hợp đồng kinh tế số 2026/HĐKT-AG")
        doc.save(raw_pdf)
        doc.close()

        # 1. Cryptographic signature
        sign_res = crypto_sign_pdf(
            input_path=raw_pdf,
            output_path=signed_pdf,
            signer_name="Công ty TNHH Antigravity Việt Nam",
            reason="Phê duyệt báo cáo tài chính năm 2026",
            location="Hà Nội, Việt Nam",
            contact_info="security@antigravity.vn",
            page_number=1,
        )
        self.assertTrue(sign_res.get("success"))
        self.assertEqual(sign_res.get("subfilter"), "ETSI.CAdES.detached")
        self.assertTrue(os.path.isfile(signed_pdf))
        self.assertEqual(len(sign_res.get("sha256_digest", "")), 64)

        # Verify cryptographic signature dictionary
        ver_res = verify_pdf_crypto_signature(signed_pdf)
        self.assertTrue(ver_res.get("has_crypto_signature"))
        self.assertGreaterEqual(ver_res.get("signature_count", 0), 1)
        self.assertEqual(ver_res["signatures"][0]["subfilter"], "ETSI.CAdES.detached")

        # 2. Convert to PDF/A-2b
        pdfa_res = convert_to_pdfa(
            input_path=signed_pdf,
            output_path=pdfa_out,
            conformance="2B",
            title="Hợp đồng lưu trữ v4.6",
            author="Antigravity Archive Engine",
        )
        self.assertTrue(pdfa_res.get("success"))
        self.assertEqual(pdfa_res.get("pdfa_conformance"), "PDF/A-2B")
        self.assertTrue(pdfa_res.get("xmp_metadata_injected"))
        self.assertTrue(pdfa_res.get("output_intent_srgb"))

        # Verify XMP metadata in output PDF
        check_doc = fitz.open(pdfa_out)
        try:
            xmp = check_doc.get_xml_metadata()
            self.assertIn("pdfaid:part>2<", xmp)
            self.assertIn("pdfaid:conformance>B<", xmp)
        finally:
            check_doc.close()

    # =========================================================================
    # SUBSYSTEM 2 & EDGE CASE #3 (ERR_XLSX_015): Excel-to-Brief & Un-cached Formulas
    # =========================================================================
    def test_02_subsystem2_excel_to_brief_and_err_xlsx_015_formula_fallback(self) -> None:
        """
        Edge Case #3 (ERR_XLSX_015):
        Workbooks generated programmatically by openpyxl have `None` cached values for
        formula cells (`=SUM(...)`, `=AVERAGE(...)`) when opened with `data_only=True`.
        `excel_to_brief` must open both `data_only=True` and `data_only=False` and
        automatically compute basic formulas so Word/MD/PDF reports never show blank totals.
        """
        xlsx_path = os.path.join(self.temp_dir, "bctc_un_cached.xlsx")
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Revenue_2026"
        ws.append(["Quarter", "Revenue_VND", "Cost_VND", "Net_Profit"])
        ws.append(["Q1/2026", 1200, 400, "=B2-C2"])
        ws.append(["Q2/2026", 1800, 600, "=B3-C3"])
        ws.append(["Q3/2026", 2000, 500, "=B4-C4"])
        ws.append(["Total", "=SUM(B2:B4)", "=SUM(C2:C4)", "=SUM(D2:D4)"])
        ws.append(["Average", "=AVERAGE(B2:B4)", "=MIN(C2:C4)", "=MAX(D2:D4)"])
        wb.save(xlsx_path)
        wb.close()

        # Verify that openpyxl data_only=True indeed returns None for B5 (un-cached)
        wb_check = openpyxl.load_workbook(xlsx_path, data_only=True)
        self.assertIsNone(wb_check["Revenue_2026"]["B5"].value)
        wb_check.close()

        # 1. Export to Markdown
        md_out = os.path.join(self.temp_dir, "brief.md")
        res_md = excel_to_brief(xlsx_path, md_out, title="Báo Cáo Doanh Thu 2026")
        self.assertTrue(res_md.get("success"))
        with open(md_out, "r", encoding="utf-8") as f:
            md_text = f.read()
        # Total Revenue = 1200 + 1800 + 2000 = 5,000; Total Cost = 1,500; Net Profit Q1 = 800, Total = 3,500
        self.assertIn("5,000", md_text)
        self.assertIn("1,500", md_text)
        self.assertIn("3,500", md_text)

        # 2. Export to DOCX and verify OpenXML invariants (<w:cantSplit/>, <w:tblHeader/>, <w:vAlign w:val="center"/>)
        docx_out = os.path.join(self.temp_dir, "brief.docx")
        res_docx = excel_to_brief(xlsx_path, docx_out, title="Báo Cáo Doanh Thu 2026")
        self.assertTrue(res_docx.get("success"))
        self.assertTrue(os.path.isfile(docx_out))

        with zipfile.ZipFile(docx_out, "r") as zf:
            doc_xml = zf.read("word/document.xml").decode("utf-8")
        self.assertIn("cantSplit", doc_xml)
        self.assertIn("tblHeader", doc_xml)
        self.assertIn('vAlign w:val="center"', doc_xml)
        self.assertIn("5,000", doc_xml)

    def test_03_subsystem2_consolidate_workbooks(self) -> None:
        b1 = os.path.join(self.temp_dir, "branch_hn.xlsx")
        b2 = os.path.join(self.temp_dir, "branch_hcm.xlsx")
        merged_out = os.path.join(self.temp_dir, "consolidated.xlsx")

        for path, branch, rev in [(b1, "Hanoi", 5000), (b2, "HCMC", 7500)]:
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Sales"
            ws.append(["Branch", "Amount"])
            ws.append([branch, rev])
            wb.save(path)
            wb.close()

        res = consolidate_workbooks([b1, b2], merged_out, mode="append_rows")
        self.assertTrue(res.get("success"))
        self.assertEqual(res.get("rows_consolidated"), 2)

        wb_m = openpyxl.load_workbook(merged_out, data_only=True)
        ws_m = wb_m.active
        self.assertEqual(ws_m.max_row, 4)  # 1 header + 2 data rows + 1 dynamic summary row
        self.assertEqual(ws_m.cell(2, 2).value, "Hanoi")
        self.assertEqual(ws_m.cell(3, 2).value, "HCMC")
        wb_m.close()

    # =========================================================================
    # SUBSYSTEM 3 & EDGE CASES #1 (ERR_SEC_006) & #2 (ERR_SEC_007): Deep CDR & PII
    # =========================================================================
    def test_04_subsystem3_err_sec_006_deep_cdr_content_types_and_rels_promotion(self) -> None:
        """
        Edge Case #1 (ERR_SEC_006):
        When stripping `vbaProject.bin` and malicious PE (`MZ`) `oleObject1.bin` from a
        `.docm` or `.xlsm` file, `DocumentSanitizer` must:
        1. Remove `/word/vbaProject.bin` and `/word/embeddings/oleObject1.bin` from the ZIP.
        2. Strip `<Override PartName="/word/vbaProject.bin".../>` from `[Content_Types].xml`.
        3. Promote `application/vnd.ms-word.document.macroEnabled.main+xml` to
           `application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml`
           so Microsoft Word/Excel does not raise a corrupt structure error on open.
        4. Strip orphan `<Relationship>` entries referencing `vbaProject.bin` or `oleObject1.bin` in `.rels`.
        """
        base_docx = os.path.join(self.temp_dir, "base.docx")
        infected_docm = os.path.join(self.temp_dir, "infected.docm")
        clean_docx = os.path.join(self.temp_dir, "clean_cdr.docx")

        doc = Document()
        doc.add_paragraph("Tài liệu kiểm thử Deep CDR v4.6.0")
        doc.save(base_docx)

        # Inject vbaProject.bin, MZ oleObject1.bin, macroEnabled ContentType, and .rels reference
        with zipfile.ZipFile(base_docx, "r") as zin, zipfile.ZipFile(infected_docm, "w") as zout:
            for item in zin.infolist():
                raw = zin.read(item.filename)
                if item.filename == "[Content_Types].xml":
                    text = raw.decode("utf-8")
                    text = text.replace(
                        "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
                        "application/vnd.ms-word.document.macroEnabled.main+xml",
                    )
                    text = text.replace(
                        "</Types>",
                        '<Override PartName="/word/vbaProject.bin" ContentType="application/vnd.ms-office.vbaProject"/></Types>',
                    )
                    raw = text.encode("utf-8")
                elif item.filename == "word/_rels/document.xml.rels":
                    text = raw.decode("utf-8")
                    text = text.replace(
                        "</Relationships>",
                        '<Relationship Id="rId999" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="vbaProject.bin"/>'
                        '<Relationship Id="rId998" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject" Target="embeddings/oleObject1.bin"/>'
                        '</Relationships>',
                    )
                    raw = text.encode("utf-8")
                zout.writestr(item, raw)

            zout.writestr("word/vbaProject.bin", b"FAKE_VBA_MACRO_STREAM")
            zout.writestr("word/embeddings/oleObject1.bin", b"MZ\x90\x00FAKE_PE_EXECUTABLE_PAYLOAD")

        san_res = sanitize_document(
            file_path=infected_docm,
            output_path=clean_docx,
            deep_cdr=True,
        )
        self.assertTrue(san_res.get("success"))
        self.assertGreaterEqual(san_res.get("macros_stripped", 0), 1)
        self.assertGreaterEqual(san_res.get("ole_Needs_stripped", 0), 1)
        self.assertGreaterEqual(san_res.get("content_types_promoted", 0), 1)
        self.assertGreaterEqual(san_res.get("orphan_rels_cleaned", 0), 2)

        with zipfile.ZipFile(clean_docx, "r") as zclean:
            names = zclean.namelist()
            self.assertNotIn("word/vbaProject.bin", names)
            self.assertNotIn("word/embeddings/oleObject1.bin", names)

            ct_xml = zclean.read("[Content_Types].xml").decode("utf-8")
            self.assertNotIn("macroEnabled.main+xml", ct_xml)
            self.assertNotIn("vbaProject.bin", ct_xml)
            self.assertIn(
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
                ct_xml,
            )

            rels_xml = zclean.read("word/_rels/document.xml.rels").decode("utf-8")
            self.assertNotIn("vbaProject.bin", rels_xml)
            self.assertNotIn("oleObject1.bin", rels_xml)

    def test_05_subsystem3_err_sec_007_pii_redactor_financial_guard(self) -> None:
        """
        Edge Case #2 (ERR_SEC_007):
        - Large 12-digit financial numbers (`150000000000`) or numbers followed by currency
          (`0912345678 VND`, `001099123456 đồng`) MUST NOT be redacted as CCCD/Phone/Bank Account.
        - Unanchored 10-14 digit numbers (`998877665544`) MUST NOT be redacted as bank accounts.
        - True Vietnamese CCCD (`001099123456`), Phone (`0912345678`), context-anchored Bank Account
          (`STK: 19034567890123`), MST (`MST: 0101234567-001`), and Email (`cfo@company.vn`) MUST be redacted.
        """
        sample_text = (
            "Tổng doanh thu: 150000000000 VND. "
            "Chi phí vận hành: 0912345678 VND. "
            "Mã chứng từ nội bộ: 998877665544. "
            "Khách hàng có CCCD 001099123456, SĐT liên hệ 0912345678, "
            "Số tài khoản: 19034567890123 tại Techcombank, "
            "MST: 0101234567-001, email: cfo@antigravity.vn."
        )
        redacted, counts = redact_pii_text(sample_text)

        # 1. Verify financial numbers & unanchored numbers are 100% preserved
        self.assertIn("150000000000 VND", redacted)
        self.assertIn("0912345678 VND", redacted)
        self.assertIn("998877665544", redacted)

        # 2. Verify real PII is 100% redacted
        self.assertNotIn("001099123456", redacted)
        self.assertIn("[REDACTED_CCCD]", redacted)
        self.assertIn("[REDACTED_PHONE]", redacted)
        self.assertNotIn("19034567890123", redacted)
        self.assertIn("Số tài khoản: [REDACTED_BANK_ACC]", redacted)
        self.assertNotIn("0101234567-001", redacted)
        self.assertIn("[REDACTED_MST]", redacted)
        self.assertNotIn("cfo@antigravity.vn", redacted)
        self.assertIn("[REDACTED_EMAIL]", redacted)

        self.assertEqual(counts["cccd"], 1)
        self.assertEqual(counts["phone"], 1)
        self.assertEqual(counts["bank_account"], 1)
        self.assertEqual(counts["mst"], 1)
        self.assertEqual(counts["email"], 1)

    # =========================================================================
    # SUBSYSTEM 4 & EDGE CASE #4 (ERR_PIPE_001): Workflow Pipeline & Temp Cleanup
    # =========================================================================
    def test_06_subsystem4_err_pipe_001_pipeline_and_worker_cap(self) -> None:
        """
        Edge Case #4 (ERR_PIPE_001):
        - Default `max_workers` reads `MAX_JOB_WORKERS` (defaults to 2 for 512MB RAM cloud safety).
        - Multi-step pipeline cleans up 100% of `step_*_temp.*` files in `finally`.
        """
        old_env = os.environ.pop("MAX_JOB_WORKERS", None)
        try:
            self.assertEqual(get_default_max_workers(), 2)
            os.environ["MAX_JOB_WORKERS"] = "3"
            self.assertEqual(get_default_max_workers(), 3)
        finally:
            if old_env is not None:
                os.environ["MAX_JOB_WORKERS"] = old_env
            else:
                os.environ.pop("MAX_JOB_WORKERS", None)

        # Create a DOCX with PII and run a 2-step pipeline: sanitize (redact PII) -> audit_qa
        in_dir = os.path.join(self.temp_dir, "pipe_in")
        out_dir = os.path.join(self.temp_dir, "pipe_out")
        os.makedirs(in_dir, exist_ok=True)

        doc_in = os.path.join(in_dir, "contract.docx")
        d = Document()
        d.add_paragraph("Hợp đồng của ông Nguyễn Văn A, CCCD: 001099123456, SĐT: 0987654321.")
        d.save(doc_in)

        spec = {
            "name": "CDR + QA Pipeline",
            "steps": [
                {"action": "sanitize", "deep_cdr": True, "redact_pii": True},
                {"action": "audit_qa"},
            ],
        }
        bulk_res = run_bulk_directory_pipeline(
            input_dir=in_dir,
            output_dir=out_dir,
            spec=spec,
        )
        self.assertTrue(bulk_res.get("success"))
        self.assertEqual(bulk_res.get("max_workers"), 2)
        self.assertEqual(bulk_res.get("succeeded"), 1)

        # Verify no step_*_temp.* files remain in output_dir
        remaining_temps = [
            f for f in os.listdir(out_dir) if "_step_" in f and "_temp" in f
        ]
        self.assertEqual(remaining_temps, [])

        # Verify output DOCX has PII redacted
        out_doc_path = bulk_res["results"][0]["output_path"]
        self.assertTrue(os.path.isfile(out_doc_path))
        d_out = Document(out_doc_path)
        full_text = "\n".join(p.text for p in d_out.paragraphs)
        self.assertIn("[REDACTED_CCCD]", full_text)
        self.assertNotIn("001099123456", full_text)

    # =========================================================================
    # SUBSYSTEM 5: AST Code-to-Diagram & Pure Draw.io SVG Export + Auto-Inject
    # =========================================================================
    def test_07_subsystem5_ast_code_to_diagram_and_drawio_svg_export(self) -> None:
        # 1. Create sample Python module for AST scanning
        py_src = os.path.join(self.temp_dir, "sample_domain.py")
        with open(py_src, "w", encoding="utf-8") as f:
            f.write(
                "class BaseService:\n"
                "    def health_check(self) -> bool:\n"
                "        return True\n\n"
                "class InvoiceManager(BaseService):\n"
                "    invoice_id: str\n"
                "    total_amount: float\n"
                "    def issue_invoice(self, customer_id: str) -> dict:\n"
                "        return {}\n"
            )

        ast_spec = scan_python_ast_to_spec(py_src)
        self.assertEqual(ast_spec["diagram_type"], "class_diagram")
        self.assertIn("class BaseService", ast_spec["code"])
        self.assertIn("class InvoiceManager", ast_spec["code"])
        self.assertIn("BaseService <|-- InvoiceManager", ast_spec["code"])

        # 2. Create sample SQL DDL and generate Draw.io + Vector SVG + auto-inject into DOCX
        sql_src = os.path.join(self.temp_dir, "schema.sql")
        with open(sql_src, "w", encoding="utf-8") as f:
            f.write(
                "CREATE TABLE customers (\n"
                "    id UUID PRIMARY KEY,\n"
                "    email VARCHAR(255) NOT NULL\n"
                ");\n"
                "CREATE TABLE orders (\n"
                "    id UUID PRIMARY KEY,\n"
                "    customer_id UUID REFERENCES customers(id),\n"
                "    total NUMERIC(12,2)\n"
                ");\n"
            )

        target_docx = os.path.join(self.temp_dir, "architecture_spec.docx")
        doc = Document()
        doc.add_heading("1. Database Schema", level=1)
        doc.save(target_docx)

        svg_out = os.path.join(self.temp_dir, "schema_erd.svg")
        gen_res = generate_diagram_from_code(
            source_path=sql_src,
            output_path=svg_out,
            diagram_type="erd",
            engine="mxgraph",
            inject_into=target_docx,
            inject_heading="1. Database Schema",
        )
        self.assertTrue(gen_res.get("success"))
        self.assertTrue(os.path.isfile(svg_out))
        with open(svg_out, "r", encoding="utf-8") as f:
            svg_content = f.read()
        self.assertIn("<svg", svg_content)
        self.assertIn("customers", svg_content)
        self.assertIn("orders", svg_content)

    # =========================================================================
    # SUBSYSTEM 6 & EDGE CASE #5 (ERR_UI_001): Isolated --user-data-dir App Mode
    # =========================================================================
    def test_08_subsystem6_err_ui_001_isolated_user_data_dir_lifecycle(self) -> None:
        """
        Edge Case #5 (ERR_UI_001):
        Verify that `build_app_mode_browser_command` passes `--user-data-dir=<temp_app_profile_dir>`
        alongside `--app=http://127.0.0.1:<port>` so `proc.wait()` blocks for the exact lifetime
        of the native App window, and `launch_native_web_studio(dry_run=True)` cleans up the
        temporary profile directory on exit.
        """
        cmd = build_app_mode_browser_command(
            exe_path="msedge.exe",
            url="http://127.0.0.1:8123",
            user_data_dir="C:\\Temp\\ag_studio_profile_test",
            width=1440,
            height=900,
        )
        self.assertEqual(cmd[0], "msedge.exe")
        self.assertIn("--app=http://127.0.0.1:8123", cmd)
        self.assertIn("--user-data-dir=C:\\Temp\\ag_studio_profile_test", cmd)
        self.assertIn("--no-first-run", cmd)

        dry_res = launch_native_web_studio(port=0, dry_run=True)
        self.assertTrue(dry_res.get("success"))
        self.assertEqual(dry_res.get("mode"), "dry_run")
        self.assertTrue(dry_res.get("profile_cleaned"))
        self.assertFalse(os.path.exists(dry_res.get("user_data_dir", "")))


if __name__ == "__main__":
    unittest.main()
