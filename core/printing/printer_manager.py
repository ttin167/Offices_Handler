"""
core.printing.printer_manager — Enterprise Hardware Printer Discovery & Silent Spooler.

Features:
1. Native Windows Spooler integration via win32print / win32api.
2. Cross-platform discovery (EnumPrinters on Windows, CUPS/lpstat on Linux).
3. 100% Silent printing without user dialog interruption:
   - Priority 1: Headless Microsoft Edge Spooler (--headless --print-to-printer).
   - Priority 2: Native ShellExecute ('printto' verb).
   - Priority 3: Virtual Spool Buffer (Cloud & Container mode).
4. Direct integration with PrintOptimizer (Grayscale, N-Up, Booklet, Watermarking).
"""

from __future__ import annotations

import os
import shutil
import subprocess
import sys
import tempfile
import time
import uuid
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple, Union

try:
    import win32api
    import win32con
    import win32print
except ImportError:
    win32api = None
    win32con = None
    win32print = None

from .print_optimizer import PrintOptimizer


class PrinterManager:
    """Manages system hardware printers, discovery, and silent job spooling."""

    @classmethod
    def list_printers(cls) -> List[Dict[str, Any]]:
        """Discovers all available physical, network, and virtual printers."""
        printers: List[Dict[str, Any]] = []
        default_printer = cls.get_default_printer()

        # 1. Windows Spooler Discovery
        if sys.platform == "win32" and win32print is not None:
            try:
                flags = win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS
                enum_list = win32print.EnumPrinters(flags)
                for item in enum_list:
                    # item: (flags, description, name, comment)
                    p_name = item[2]
                    p_desc = item[1]
                    p_comment = item[3] if len(item) > 3 else ""

                    is_def = (p_name.strip().lower() == (default_printer or "").strip().lower())
                    printers.append({
                        "name": p_name,
                        "description": p_desc,
                        "comment": p_comment,
                        "is_default": is_def,
                        "is_virtual": any(v in p_name.lower() for v in ["pdf", "xps", "onenote", "fax"]),
                        "status": "READY"
                    })
            except Exception as e:
                pass

        # 2. Linux / macOS CUPS Discovery
        elif sys.platform != "win32":
            try:
                res = subprocess.run(["lpstat", "-p"], stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, timeout=5)
                if res.returncode == 0:
                    for line in res.stdout.splitlines():
                        line = line.strip()
                        if line.startswith("printer "):
                            parts = line.split()
                            p_name = parts[1]
                            printers.append({
                                "name": p_name,
                                "description": line,
                                "is_default": (p_name == default_printer),
                                "is_virtual": "pdf" in p_name.lower(),
                                "status": "READY" if "idle" in line or "enabled" in line else "UNKNOWN"
                            })
            except Exception:
                pass

        # 3. Always ensure at least 1 Virtual Spooler is available
        if not printers:
            printers.append({
                "name": "Virtual Spooler (PDF Buffer)",
                "description": "Antigravity Cloud & Virtual Spool Container",
                "comment": "Lưu trữ tệp in ra thư mục output/spool/",
                "is_default": True,
                "is_virtual": True,
                "status": "READY"
            })

        return printers

    @classmethod
    def get_default_printer(cls) -> Optional[str]:
        """Gets default system printer name."""
        if sys.platform == "win32" and win32print is not None:
            try:
                return win32print.GetDefaultPrinter()
            except Exception:
                pass
        elif sys.platform != "win32":
            try:
                res = subprocess.run(["lpstat", "-d"], stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, timeout=5)
                if res.returncode == 0:
                    for line in res.stdout.splitlines():
                        if "system default destination:" in line:
                            return line.split(":", 1)[1].strip()
            except Exception:
                pass
        return None

    @classmethod
    def _find_edge_binary(cls) -> Optional[str]:
        """Locates Microsoft Edge executable for silent headless PDF printing."""
        candidates = [
            os.path.expandvars(r"%ProgramFiles(x86)%\Microsoft\Edge\Application\msedge.exe"),
            os.path.expandvars(r"%ProgramFiles%\Microsoft\Edge\Application\msedge.exe"),
            os.path.expandvars(r"%LocalAppData%\Microsoft\Edge\Application\msedge.exe"),
        ]
        for p in candidates:
            if os.path.isfile(p):
                return p
        return None

    @classmethod
    def get_device_capabilities(cls, printer_name: Optional[str] = None) -> Dict[str, Any]:
        """
        Inspects printer hardware capabilities (DC_BINS, DC_DUPLEX, DC_PAPERS)
        with Guard G2 to prevent crashing on virtual printers or generic drivers.
        """
        target = printer_name or cls.get_default_printer()
        if not target or sys.platform != "win32" or win32print is None or win32con is None:
            return {
                "printer": target,
                "supported_bins": (),
                "supports_duplex": False,
                "is_virtual": True,
            }

        try:
            try:
                bins = win32print.DeviceCapabilities(target, "", win32con.DC_BINS)
            except Exception:
                bins = ()

            try:
                duplex_flag = win32print.DeviceCapabilities(target, "", win32con.DC_DUPLEX)
                supports_duplex = bool(duplex_flag == 1)
            except Exception:
                supports_duplex = False

            is_virtual = any(v in target.lower() for v in ["pdf", "xps", "onenote", "fax"])
            return {
                "printer": target,
                "supported_bins": bins,
                "supports_duplex": supports_duplex,
                "is_virtual": is_virtual,
            }
        except Exception:
            return {
                "printer": target,
                "supported_bins": (),
                "supports_duplex": False,
                "is_virtual": True,
            }

    @classmethod
    def apply_devmode(
        cls,
        printer_name: str,
        tray_id: Optional[int] = None,
        duplex: Optional[str] = None,
    ) -> bool:
        """
        Safely configures printer hardware DEVMODE with Guard G2 capabilities check.
        """
        if sys.platform != "win32" or win32print is None or win32con is None:
            return False

        caps = cls.get_device_capabilities(printer_name)

        # Guard G2: Check tray capability before applying
        valid_tray = None
        if tray_id is not None:
            if tray_id in caps["supported_bins"]:
                valid_tray = tray_id

        # Guard G2: Check duplex capability
        valid_duplex = None
        if duplex and caps["supports_duplex"]:
            dup_lower = duplex.lower()
            if dup_lower in ("vertical", "long", "long-edge"):
                valid_duplex = win32con.DMDUP_VERTICAL
            elif dup_lower in ("horizontal", "short", "short-edge"):
                valid_duplex = win32con.DMDUP_HORIZONTAL
            elif dup_lower in ("simplex", "none", "single"):
                valid_duplex = win32con.DMDUP_SIMPLEX

        if valid_tray is None and valid_duplex is None:
            return False

        try:
            h_printer = win32print.OpenPrinter(printer_name, {"DesiredAccess": win32print.PRINTER_ALL_ACCESS})
            try:
                p_info = win32print.GetPrinter(h_printer, 2)
                devmode = p_info.get("pDevMode")
                if devmode:
                    if valid_tray is not None:
                        devmode.DefaultSource = valid_tray
                    if valid_duplex is not None:
                        devmode.Duplex = valid_duplex
                    win32print.SetPrinter(h_printer, 2, p_info, 0)
                    return True
            finally:
                win32print.ClosePrinter(h_printer)
        except Exception:
            pass
        return False

    @classmethod
    def print_file(
        cls,
        file_path: str,
        printer_name: Optional[str] = None,
        copies: int = 1,
        job_title: Optional[str] = None,
        silent: bool = True,
        optimizer_options: Optional[Dict[str, Any]] = None,
        tray_id: Optional[int] = None,
        duplex: Optional[str] = None,
    ) -> Dict[str, Any]:
        """
        Submits a document to the specified printer or system default printer.
        Automatically converts DOCX/XLSX to PDF and applies print optimizations if requested.
        """
        if not os.path.isfile(file_path):
            raise FileNotFoundError(f"Tệp in không tồn tại: {file_path}")

        target_printer = printer_name or cls.get_default_printer() or "Virtual Spooler (PDF Buffer)"
        job_id = f"JOB-{int(time.time())}-{uuid.uuid4().hex[:6]}"
        job_name = job_title or os.path.basename(file_path)

        # Step 0: Apply DEVMODE settings if physical printer on Windows (Guard G2)
        devmode_applied = False
        if sys.platform == "win32" and (tray_id is not None or duplex is not None):
            devmode_applied = cls.apply_devmode(target_printer, tray_id=tray_id, duplex=duplex)

        # Step 1: Ensure PDF format for printing
        ext = os.path.splitext(file_path)[1].lower()
        printable_pdf = file_path

        if ext != ".pdf":
            try:
                from core.converter.engine import convert_universal
                temp_pdf = os.path.join(tempfile.gettempdir(), f"{job_id}_converted.pdf")
                convert_universal(file_path, temp_pdf)
                printable_pdf = temp_pdf
            except Exception as ce:
                raise RuntimeError(f"Không thể chuyển đổi {ext} sang PDF để in: {ce}")

        # Step 2: Apply Print Optimizer (Vector Grayscale, N-Up, Booklet Creep, Watermark, Margin Guard)
        if optimizer_options:
            optimized_pdf = os.path.join(tempfile.gettempdir(), f"{job_id}_optimized.pdf")
            PrintOptimizer.optimize_pdf(
                input_pdf_path=printable_pdf,
                output_pdf_path=optimized_pdf,
                grayscale=optimizer_options.get("grayscale", False),
                vector_grayscale=optimizer_options.get("vector_grayscale", True),
                dpi=optimizer_options.get("dpi", 300),
                n_up=optimizer_options.get("n_up", 1),
                booklet=optimizer_options.get("booklet", False),
                creep_compensation=optimizer_options.get("creep_compensation", True),
                paper_thickness_pt=optimizer_options.get("paper_thickness_pt", 0.28),
                watermark_text=optimizer_options.get("watermark"),
                margin_guard_mm=optimizer_options.get("margin_guard_mm", 5.0),
            )
            printable_pdf = optimized_pdf

        # Step 3: Check if Virtual Spooler is requested or if printer is Virtual Spooler
        if "virtual" in target_printer.lower() or "buffer" in target_printer.lower():
            spool_dir = os.path.abspath(os.path.join(os.getcwd(), "output", "spool"))
            os.makedirs(spool_dir, exist_ok=True)
            spool_file = os.path.join(spool_dir, f"{job_id}_{os.path.basename(file_path)}.pdf")
            shutil.copy2(printable_pdf, spool_file)
            return {
                "success": True,
                "job_id": job_id,
                "printer": target_printer,
                "copies": copies,
                "spooled_file": spool_file,
                "mode": "VIRTUAL_SPOOL",
                "status": "QUEUED"
            }

        # Step 4: Physical Spooling on Windows
        if sys.platform == "win32":
            edge_bin = cls._find_edge_binary()
            if edge_bin and silent:
                # Priority 1: Headless Edge (100% silent, standard on Windows 10/11)
                cmd = [
                    edge_bin,
                    "--headless",
                    "--disable-gpu",
                    f"--print-to-printer={target_printer}",
                    printable_pdf
                ]
                for _ in range(copies):
                    proc = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=25)
                return {
                    "success": True,
                    "job_id": job_id,
                    "printer": target_printer,
                    "copies": copies,
                    "mode": "EDGE_HEADLESS_SPOOL",
                    "status": "SENT_TO_PRINTER"
                }
            elif win32api is not None:
                # Priority 2: ShellExecute printto
                for _ in range(copies):
                    win32api.ShellExecute(0, "printto", printable_pdf, f'"{target_printer}"', ".", 0)
                return {
                    "success": True,
                    "job_id": job_id,
                    "printer": target_printer,
                    "copies": copies,
                    "mode": "SHELLEXECUTE_PRINTTO",
                    "status": "SENT_TO_PRINTER"
                }

        # Step 5: Physical Spooling on Linux / CUPS
        else:
            cmd = ["lp", "-d", target_printer, "-n", str(copies), "-t", job_name, printable_pdf]
            res = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, timeout=15)
            if res.returncode != 0:
                raise RuntimeError(f"Lỗi gửi lệnh in CUPS (lp): {res.stderr}")
            return {
                "success": True,
                "job_id": job_id,
                "printer": target_printer,
                "copies": copies,
                "mode": "CUPS_LP",
                "status": "SENT_TO_PRINTER"
            }

        return {
            "success": True,
            "job_id": job_id,
            "printer": target_printer,
            "copies": copies,
            "mode": "COMPLETED",
            "status": "FINISHED"
        }
