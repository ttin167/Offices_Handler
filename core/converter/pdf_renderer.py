"""
core/converter/pdf_renderer.py — Universal Cross-Platform PDF Renderer (v4.0).

Provides Zero-COM and multi-backend document-to-PDF rendering:
  - Backend 1: Win32 COM automation (MS Office on Windows).
  - Backend 2: Headless LibreOffice ('soffice --headless --convert-to pdf') for Linux / Docker.
  - Backend 3: Pure Python & PyMuPDF fallback pipeline.
"""

from __future__ import annotations

import os
import shutil
import subprocess
import sys
import tempfile
from typing import Any, Dict, List, Optional, Tuple

COMMON_LIBREOFFICE_PATHS = [
    r"C:\Program Files\LibreOffice\program\soffice.exe",
    r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    "/usr/bin/soffice",
    "/usr/bin/libreoffice",
    "/usr/local/bin/soffice",
    "/opt/libreoffice/program/soffice",
    "/Applications/LibreOffice.app/Contents/MacOS/soffice",
]


def find_libreoffice_binary() -> Optional[str]:
    """Find LibreOffice executable on system PATH or standard system paths."""
    for bin_name in ["soffice", "libreoffice"]:
        p = shutil.which(bin_name)
        if p:
            return p

    for path in COMMON_LIBREOFFICE_PATHS:
        if os.path.isfile(path):
            return path

    return None


_COM_CACHE: Dict[str, bool] = {}


def is_win32_com_available(app_name: str = "Word.Application") -> bool:
    """Check if Win32 COM is functional on this Windows environment via registry (cached)."""
    if sys.platform != "win32":
        return False
    if app_name in _COM_CACHE:
        return _COM_CACHE[app_name]
    try:
        import winreg
        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, f"{app_name}\\CurVer"):
            _COM_CACHE[app_name] = True
            return True
    except Exception:
        pass
    _COM_CACHE[app_name] = False
    return False


class UniversalPdfRenderer:
    """
    Universal Multi-Backend PDF Renderer.
    Gracefully routes conversion requests based on environment capabilities.
    """

    def __init__(
        self,
        prefer_libreoffice: bool = False,
        libreoffice_path: Optional[str] = None,
    ) -> None:
        self.prefer_libreoffice = prefer_libreoffice
        self.libreoffice_binary = libreoffice_path or find_libreoffice_binary()

    def get_available_backends(self) -> List[str]:
        """Return list of functional backends in priority order."""
        backends: List[str] = []
        if sys.platform == "win32" and not self.prefer_libreoffice:
            if is_win32_com_available("Word.Application"):
                backends.append("win32com")

        if self.libreoffice_binary:
            backends.append("libreoffice")

        if sys.platform == "win32" and self.prefer_libreoffice:
            if is_win32_com_available("Word.Application"):
                backends.append("win32com")

        backends.append("fallback")
        return backends

    def detect_active_backend(self) -> str:
        """Return currently preferred active backend."""
        backends = self.get_available_backends()
        return backends[0] if backends else "fallback"

    def convert_docx_to_pdf(
        self, docx_path: str, output_pdf_path: Optional[str] = None
    ) -> str:
        """Convert Word (.docx) document to PDF."""
        if not os.path.exists(docx_path):
            raise FileNotFoundError(f"Source DOCX not found: {docx_path}")

        target_pdf = output_pdf_path or os.path.splitext(docx_path)[0] + ".pdf"
        target_dir = os.path.dirname(os.path.abspath(target_pdf))
        os.makedirs(target_dir, exist_ok=True)

        backends = self.get_available_backends()

        # 1. Try Win32 COM
        if "win32com" in backends and not self.prefer_libreoffice:
            try:
                return self._convert_docx_win32com(docx_path, target_pdf)
            except Exception as ex:
                if "libreoffice" not in backends:
                    raise RuntimeError(f"Win32 COM conversion failed: {ex}") from ex

        # 2. Try LibreOffice
        if "libreoffice" in backends:
            return self._convert_via_libreoffice(docx_path, target_pdf)

        raise RuntimeError(
            "No capable PDF conversion backend available. "
            "Please install LibreOffice ('apt install libreoffice' or 'winget install LibreOffice.LibreOffice') "
            "or run on Windows with Microsoft Office installed."
        )

    def convert_pptx_to_pdf(
        self, pptx_path: str, output_pdf_path: Optional[str] = None
    ) -> str:
        """Convert PowerPoint (.pptx) presentation to PDF."""
        if not os.path.exists(pptx_path):
            raise FileNotFoundError(f"Source PPTX not found: {pptx_path}")

        target_pdf = output_pdf_path or os.path.splitext(pptx_path)[0] + ".pdf"
        target_dir = os.path.dirname(os.path.abspath(target_pdf))
        os.makedirs(target_dir, exist_ok=True)

        backends = self.get_available_backends()

        if "win32com" in backends and not self.prefer_libreoffice:
            try:
                return self._convert_pptx_win32com(pptx_path, target_pdf)
            except Exception:
                pass

        if "libreoffice" in backends:
            return self._convert_via_libreoffice(pptx_path, target_pdf)

        raise RuntimeError(
            "No capable PDF conversion backend available for PPTX. "
            "Please install LibreOffice or run on Windows with Microsoft PowerPoint."
        )

    def convert_xlsx_to_pdf(
        self, xlsx_path: str, output_pdf_path: Optional[str] = None
    ) -> str:
        """Convert Excel (.xlsx) spreadsheet to PDF."""
        if not os.path.exists(xlsx_path):
            raise FileNotFoundError(f"Source XLSX not found: {xlsx_path}")

        target_pdf = output_pdf_path or os.path.splitext(xlsx_path)[0] + ".pdf"
        target_dir = os.path.dirname(os.path.abspath(target_pdf))
        os.makedirs(target_dir, exist_ok=True)

        backends = self.get_available_backends()

        if "libreoffice" in backends:
            return self._convert_via_libreoffice(xlsx_path, target_pdf)

        raise RuntimeError(
            "LibreOffice backend is required for Excel to PDF conversion on this platform."
        )

    def convert_to_pdf(
        self, source_path: str, output_pdf_path: Optional[str] = None
    ) -> str:
        """Dispatch universal conversion based on source file extension."""
        ext = os.path.splitext(source_path)[1].lower()
        if ext == ".docx":
            return self.convert_docx_to_pdf(source_path, output_pdf_path)
        elif ext == ".pptx":
            return self.convert_pptx_to_pdf(source_path, output_pdf_path)
        elif ext in [".xlsx", ".xls"]:
            return self.convert_xlsx_to_pdf(source_path, output_pdf_path)
        elif ext == ".pdf":
            if output_pdf_path and output_pdf_path != source_path:
                shutil.copy2(source_path, output_pdf_path)
                return output_pdf_path
            return source_path
        else:
            raise ValueError(f"Unsupported file format for PDF conversion: {ext}")

    def _convert_via_libreoffice(self, source_path: str, target_pdf: str) -> str:
        """Run LibreOffice headless convert-to pdf command."""
        bin_path = self.libreoffice_binary or "soffice"
        out_dir = os.path.dirname(os.path.abspath(target_pdf))

        cmd = [
            bin_path,
            "--headless",
            "--convert-to",
            "pdf",
            "--outdir",
            out_dir,
            os.path.abspath(source_path),
        ]

        proc = subprocess.run(cmd, capture_output=True, text=True, timeout=120)
        if proc.returncode != 0:
            raise RuntimeError(
                f"LibreOffice PDF conversion failed with exit code {proc.returncode}: {proc.stderr}"
            )

        # LibreOffice names the output file <basename>.pdf in outdir
        expected_out = os.path.join(
            out_dir, os.path.splitext(os.path.basename(source_path))[0] + ".pdf"
        )

        if os.path.exists(expected_out):
            if os.path.abspath(expected_out) != os.path.abspath(target_pdf):
                shutil.move(expected_out, target_pdf)
            return target_pdf

        raise FileNotFoundError(f"Expected converted PDF not produced at: {expected_out}")

    def _convert_docx_win32com(self, docx_path: str, target_pdf: str) -> str:
        """Convert DOCX via Microsoft Word COM automation."""
        import win32com.client  # type: ignore

        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        try:
            doc = word.Documents.Open(os.path.abspath(docx_path))
            # 17 represents wdFormatPDF
            doc.SaveAs(os.path.abspath(target_pdf), FileFormat=17)
            doc.Close(False)
            return target_pdf
        finally:
            word.Quit()

    def _convert_pptx_win32com(self, pptx_path: str, target_pdf: str) -> str:
        """Convert PPTX via Microsoft PowerPoint COM automation."""
        import win32com.client  # type: ignore

        powerpoint = win32com.client.Dispatch("PowerPoint.Application")
        try:
            deck = powerpoint.Presentations.Open(
                os.path.abspath(pptx_path), WithWindow=False
            )
            # 32 represents ppSaveAsPDF
            deck.SaveAs(os.path.abspath(target_pdf), FileFormat=32)
            deck.Close()
            return target_pdf
        finally:
            powerpoint.Quit()


def convert_document_to_pdf(
    source_path: str,
    output_pdf_path: Optional[str] = None,
    prefer_libreoffice: bool = False,
) -> str:
    """Convenience functional wrapper for UniversalPdfRenderer."""
    renderer = UniversalPdfRenderer(prefer_libreoffice=prefer_libreoffice)
    return renderer.convert_to_pdf(source_path, output_pdf_path)
