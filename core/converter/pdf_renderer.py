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


def _dispatch_isolated_com(prog_id: str) -> Any:
    """Create an isolated COM server instance and register its HWND PID with ComProcessWatchdog."""
    import win32com.client  # type: ignore

    app = win32com.client.DispatchEx(prog_id)
    try:
        from core.converter.com_watchdog import ComProcessWatchdog

        hwnd = getattr(app, "Hwnd", None) or getattr(app, "HWND", None)
        if hwnd:
            ComProcessWatchdog.register_pid_from_hwnd(hwnd)
    except Exception:
        pass
    return app


def _run_with_com_watchdog(func: Any, app_type: str, timeout_sec: float = 25.0) -> Any:
    try:
        from core.converter.com_watchdog import ComProcessWatchdog

        return ComProcessWatchdog.run_with_timeout(
            func, timeout_sec=timeout_sec, app_type=app_type
        )
    except (ImportError, ValueError):
        return func()


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

    @staticmethod
    def _any_office_com_available() -> bool:
        return any(
            is_win32_com_available(prog_id)
            for prog_id in ("Word.Application", "Excel.Application", "PowerPoint.Application")
        )

    def get_available_backends(self) -> List[str]:
        """Return list of functional backends in priority order."""
        backends: List[str] = []
        has_com = sys.platform == "win32" and self._any_office_com_available()
        if has_com and not self.prefer_libreoffice:
            backends.append("win32com")

        if self.libreoffice_binary:
            backends.append("libreoffice")

        if has_com and self.prefer_libreoffice:
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
        """Convert Excel (.xlsx/.xlsm/.xls) spreadsheet to PDF via Excel COM or LibreOffice."""
        if not os.path.exists(xlsx_path):
            raise FileNotFoundError(f"Source XLSX not found: {xlsx_path}")

        target_pdf = output_pdf_path or os.path.splitext(xlsx_path)[0] + ".pdf"
        target_dir = os.path.dirname(os.path.abspath(target_pdf))
        os.makedirs(target_dir, exist_ok=True)

        backends = self.get_available_backends()
        has_excel_com = sys.platform == "win32" and is_win32_com_available("Excel.Application")
        has_lo = "libreoffice" in backends

        attempts: List[str] = []
        if self.prefer_libreoffice:
            if has_lo:
                attempts.append("libreoffice")
            if has_excel_com:
                attempts.append("excel_com")
        else:
            if has_excel_com:
                attempts.append("excel_com")
            if has_lo:
                attempts.append("libreoffice")

        last_err: Optional[Exception] = None
        for backend in attempts:
            try:
                if backend == "excel_com":
                    return self._convert_xlsx_win32com(xlsx_path, target_pdf)
                return self._convert_via_libreoffice(xlsx_path, target_pdf)
            except Exception as ex:
                last_err = ex

        if last_err is not None:
            raise RuntimeError(f"Excel to PDF conversion failed: {last_err}") from last_err

        raise RuntimeError(
            "Microsoft Excel COM (Windows) or LibreOffice backend is required for Excel to PDF conversion on this platform."
        )

    def convert_xls_to_xlsx(
        self, xls_path: str, output_xlsx_path: Optional[str] = None
    ) -> str:
        """Upgrade legacy .xls binary workbook to modern OpenXML .xlsx."""
        if not os.path.exists(xls_path):
            raise FileNotFoundError(f"Source XLS not found: {xls_path}")

        target_xlsx = output_xlsx_path or os.path.splitext(xls_path)[0] + ".xlsx"
        target_dir = os.path.dirname(os.path.abspath(target_xlsx))
        os.makedirs(target_dir, exist_ok=True)

        backends = self.get_available_backends()
        has_excel_com = sys.platform == "win32" and is_win32_com_available("Excel.Application")
        has_lo = "libreoffice" in backends

        order = (
            (["libreoffice"] if has_lo else []) + (["excel_com"] if has_excel_com else [])
            if self.prefer_libreoffice
            else (["excel_com"] if has_excel_com else []) + (["libreoffice"] if has_lo else [])
        )

        last_err: Optional[Exception] = None
        for backend in order:
            try:
                if backend == "excel_com":
                    return self._convert_xls_win32com(xls_path, target_xlsx)
                return self._convert_via_libreoffice_format(xls_path, target_xlsx, "xlsx")
            except Exception as exc:
                last_err = exc

        # Optional pure-Python fallback via xlrd + openpyxl if installed
        try:
            return self._convert_xls_via_xlrd(xls_path, target_xlsx)
        except Exception as xlrd_exc:
            if last_err is not None:
                raise RuntimeError(
                    f"Legacy .xls conversion failed ({last_err}; xlrd fallback: {xlrd_exc})"
                ) from last_err
            raise RuntimeError(
                "Legacy .xls conversion requires Microsoft Excel COM, LibreOffice, or 'xlrd' package."
            ) from xlrd_exc

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

    def _convert_via_libreoffice_format(
        self, source_path: str, target_path: str, target_fmt: str
    ) -> str:
        """Run LibreOffice headless convert-to <target_fmt> command."""
        bin_path = self.libreoffice_binary or "soffice"
        out_dir = os.path.dirname(os.path.abspath(target_path))

        cmd = [
            bin_path,
            "--headless",
            "--convert-to",
            target_fmt,
            "--outdir",
            out_dir,
            os.path.abspath(source_path),
        ]

        proc = subprocess.run(cmd, capture_output=True, text=True, timeout=120)
        if proc.returncode != 0:
            raise RuntimeError(
                f"LibreOffice {target_fmt.upper()} conversion failed with exit code {proc.returncode}: {proc.stderr}"
            )

        expected_out = os.path.join(
            out_dir, f"{os.path.splitext(os.path.basename(source_path))[0]}.{target_fmt}"
        )

        if os.path.exists(expected_out):
            if os.path.abspath(expected_out) != os.path.abspath(target_path):
                shutil.move(expected_out, target_path)
            return target_path

        raise FileNotFoundError(f"Expected converted output not produced at: {expected_out}")

    def _convert_via_libreoffice(self, source_path: str, target_pdf: str) -> str:
        """Run LibreOffice headless convert-to pdf command."""
        return self._convert_via_libreoffice_format(source_path, target_pdf, "pdf")

    def _convert_docx_win32com(self, docx_path: str, target_pdf: str) -> str:
        """Convert DOCX via isolated Microsoft Word COM automation with watchdog protection."""
        def _do_convert() -> str:
            word = _dispatch_isolated_com("Word.Application")
            word.Visible = False
            try:
                word.DisplayAlerts = 0
            except Exception:
                pass
            try:
                doc = word.Documents.Open(os.path.abspath(docx_path), ReadOnly=True)
                # 17 represents wdFormatPDF
                doc.SaveAs(os.path.abspath(target_pdf), FileFormat=17)
                doc.Close(False)
                return target_pdf
            finally:
                try:
                    word.Quit()
                except Exception:
                    pass

        return _run_with_com_watchdog(_do_convert, app_type="word", timeout_sec=20.0)

    def _convert_pptx_win32com(self, pptx_path: str, target_pdf: str) -> str:
        """Convert PPTX via isolated Microsoft PowerPoint COM automation with watchdog protection."""
        def _do_convert() -> str:
            powerpoint = _dispatch_isolated_com("PowerPoint.Application")
            try:
                deck = powerpoint.Presentations.Open(
                    os.path.abspath(pptx_path), WithWindow=False
                )
                # 32 represents ppSaveAsPDF
                deck.SaveAs(os.path.abspath(target_pdf), FileFormat=32)
                deck.Close()
                return target_pdf
            finally:
                try:
                    powerpoint.Quit()
                except Exception:
                    pass

        return _run_with_com_watchdog(_do_convert, app_type="powerpoint", timeout_sec=20.0)

    def _convert_xlsx_win32com(self, xlsx_path: str, target_pdf: str) -> str:
        """Convert XLSX/XLS/XLSM via isolated Microsoft Excel COM ExportAsFixedFormat (PDF)."""
        abs_src = os.path.abspath(xlsx_path)
        abs_pdf = os.path.abspath(target_pdf)

        def _do_convert() -> str:
            excel = _dispatch_isolated_com("Excel.Application")
            excel.Visible = False
            try:
                excel.DisplayAlerts = False
            except Exception:
                pass
            wb = None
            try:
                wb = excel.Workbooks.Open(abs_src, UpdateLinks=0, ReadOnly=True)
                # 0 represents xlTypePDF
                wb.ExportAsFixedFormat(Type=0, Filename=abs_pdf)
                return target_pdf
            finally:
                if wb is not None:
                    try:
                        wb.Close(SaveChanges=False)
                    except Exception:
                        pass
                try:
                    excel.Quit()
                except Exception:
                    pass

        return _run_with_com_watchdog(_do_convert, app_type="excel", timeout_sec=25.0)

    def _convert_xls_win32com(self, xls_path: str, target_xlsx: str) -> str:
        """Convert legacy .xls via isolated Microsoft Excel COM SaveAs (FileFormat=51)."""
        abs_src = os.path.abspath(xls_path)
        abs_xlsx = os.path.abspath(target_xlsx)

        def _do_convert() -> str:
            excel = _dispatch_isolated_com("Excel.Application")
            excel.Visible = False
            try:
                excel.DisplayAlerts = False
            except Exception:
                pass
            wb = None
            try:
                wb = excel.Workbooks.Open(abs_src, UpdateLinks=0, ReadOnly=True)
                # 51 represents xlOpenXMLWorkbook (.xlsx)
                wb.SaveAs(abs_xlsx, FileFormat=51)
                return target_xlsx
            finally:
                if wb is not None:
                    try:
                        wb.Close(SaveChanges=False)
                    except Exception:
                        pass
                try:
                    excel.Quit()
                except Exception:
                    pass

        return _run_with_com_watchdog(_do_convert, app_type="excel", timeout_sec=25.0)

    @staticmethod
    def _convert_xls_via_xlrd(xls_path: str, target_xlsx: str) -> str:
        """Pure-Python legacy .xls -> .xlsx converter via optional xlrd + openpyxl."""
        import openpyxl  # type: ignore
        import xlrd  # type: ignore

        book = xlrd.open_workbook(os.path.abspath(xls_path))
        wb_out = openpyxl.Workbook()
        # Remove initial blank sheet after writing converted sheets
        first_sheet = wb_out.active
        wrote_any = False

        for s_idx in range(book.nsheets):
            sh = book.sheet_by_index(s_idx)
            ws = first_sheet if not wrote_any else wb_out.create_sheet(title=sh.name[:31] or f"Sheet{s_idx+1}")
            if not wrote_any:
                ws.title = sh.name[:31] or "Sheet1"
                wrote_any = True
            for r in range(sh.nrows):
                for c in range(sh.ncols):
                    cell = sh.cell(r, c)
                    val = cell.value
                    if cell.ctype == xlrd.XL_CELL_DATE:
                        val = xlrd.xldate_as_datetime(val, book.datemode)
                    elif cell.ctype == xlrd.XL_CELL_BOOLEAN:
                        val = bool(val)
                    elif cell.ctype == xlrd.XL_CELL_EMPTY:
                        val = None
                    ws.cell(row=r + 1, column=c + 1, value=val)

        wb_out.save(os.path.abspath(target_xlsx))
        wb_out.close()
        return target_xlsx


def convert_xls_to_xlsx(
    xls_path: str,
    output_xlsx_path: Optional[str] = None,
    prefer_libreoffice: bool = False,
) -> str:
    """Convenience functional wrapper to upgrade legacy .xls to .xlsx."""
    renderer = UniversalPdfRenderer(prefer_libreoffice=prefer_libreoffice)
    return renderer.convert_xls_to_xlsx(xls_path, output_xlsx_path)


def convert_document_to_pdf(
    source_path: str,
    output_pdf_path: Optional[str] = None,
    prefer_libreoffice: bool = False,
) -> str:
    """Convenience functional wrapper for UniversalPdfRenderer."""
    renderer = UniversalPdfRenderer(prefer_libreoffice=prefer_libreoffice)
    return renderer.convert_to_pdf(source_path, output_pdf_path)
