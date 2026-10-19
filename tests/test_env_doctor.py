"""
tests/test_env_doctor.py — Unit Tests for Zero-Setup Environment Doctor.

Tests:
1. EnvironmentDoctor initialization and path discovery.
2. Python dependency scanning (python-docx, pptx, openpyxl, pymupdf, Pillow).
3. Java & plantuml.jar binary discovery.
4. Chromium/Edge headless runner detection.
5. Word COM / LibreOffice document exporter validation.
6. Tesseract OCR scanner.
7. Aggregated diagnostic report generation and formatting.
8. Root facade re-export parity.
"""

from __future__ import annotations

import unittest
from core.env_doctor import EnvironmentDoctor, run_doctor, format_doctor_report
import env_doctor


class TestEnvironmentDoctor(unittest.TestCase):
    def setUp(self):
        self.doctor = EnvironmentDoctor()

    def test_01_init_and_workspace_root(self):
        self.assertIsNotNone(self.doctor.workspace_root)
        self.assertTrue(len(self.doctor.workspace_root) > 0)

    def test_02_check_python_dependencies(self):
        res = self.doctor.check_python_dependencies()
        self.assertIn("status", res)
        self.assertIn("installed", res)
        self.assertIn("missing", res)
        # In current virtualenv, required core dependencies should be installed
        self.assertIn("python-docx", res["installed"])
        self.assertIn("python-pptx", res["installed"])
        self.assertIn("openpyxl", res["installed"])
        self.assertEqual(res["status"], "healthy")

    def test_03_check_java_plantuml(self):
        res = self.doctor.check_java_plantuml()
        self.assertIn("status", res)
        self.assertIn("has_java", res)
        self.assertIn("jar_path", res)
        self.assertIn("jar_size_mb", res)

    def test_04_check_chromium_runner(self):
        res = self.doctor.check_chromium_runner()
        self.assertIn("status", res)
        self.assertIn("found", res)
        self.assertIn("path", res)

    def test_05_check_office_converters(self):
        res = self.doctor.check_office_converters()
        self.assertIn("status", res)
        self.assertIn("word_com_available", res)
        self.assertIn("libreoffice_path", res)

    def test_06_check_ocr_tesseract(self):
        res = self.doctor.check_ocr_tesseract()
        self.assertIn("found", res)
        self.assertIn("path", res)
        self.assertIn("has_vietnamese", res)

    def test_07_run_diagnostics_aggregation(self):
        report = run_doctor()
        self.assertIn("status", report)
        self.assertIn(report["status"], ("healthy", "degraded", "critical"))
        self.assertIn("platform", report)
        self.assertIn("python_version", report)
        self.assertIn("components", report)
        self.assertIn("recommendations", report)
        self.assertIn("elapsed_sec", report)
        self.assertGreater(report["elapsed_sec"], 0)

    def test_08_format_doctor_report(self):
        report = run_doctor()
        formatted = format_doctor_report(report)
        self.assertIsInstance(formatted, str)
        self.assertIn("ANTIGRAVITY OFFICE ENVIRONMENT DOCTOR", formatted)
        self.assertIn("Python Libraries", formatted)

    def test_09_root_facade_parity(self):
        self.assertEqual(env_doctor.run_doctor, run_doctor)
        self.assertEqual(env_doctor.format_doctor_report, format_doctor_report)
        self.assertEqual(env_doctor.EnvironmentDoctor, EnvironmentDoctor)


if __name__ == "__main__":
    unittest.main()
