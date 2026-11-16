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
    import win32gui
    import win32print
    import win32ui
except ImportError:
    win32api = None
    win32con = None
    win32gui = None
    win32print = None
    win32ui = None

try:
    from PIL import Image, ImageWin
except ImportError:
    Image = None
    ImageWin = None

try:
    import pymupdf as fitz
except ImportError:
    try:
        import fitz  # type: ignore
    except ImportError:
        fitz = None

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
    def _spool_via_gdi(
        cls,
        pdf_path: str,
        printer_name: str,
        job_name: str = "Antigravity Print Job",
        copies: int = 1,
        target_dpi: int = 600,
        devmode: Any = None,
        grayscale: bool = False,
    ) -> bool:
        """
        Natively spools PDF pages directly into Windows GDI printer DC with optional
        per-job DEVMODE (tray/duplex without global SetPrinter mutation) and 1-channel
        grayscale optimization for 3x lighter monochrome spooler payloads.
        """
        if sys.platform != "win32" or win32ui is None or win32con is None or Image is None or ImageWin is None or fitz is None:
            return False

        if not os.path.isfile(pdf_path):
            return False

        doc = fitz.open(pdf_path)
        if len(doc) == 0:
            doc.close()
            return False

        hdc = None
        doc_started = False
        try:
            if devmode is not None and win32gui is not None:
                raw_hdc = win32gui.CreateDC("WINSPOOL", printer_name, devmode)
                if not raw_hdc:
                    return False
                hdc = win32ui.CreateDCFromHandle(raw_hdc)
            else:
                hdc = win32ui.CreateDC()
                hdc.CreatePrinterDC(printer_name)

            horz_res = hdc.GetDeviceCaps(win32con.HORZRES)
            vert_res = hdc.GetDeviceCaps(win32con.VERTRES)
            log_x = hdc.GetDeviceCaps(win32con.LOGPIXELSX) or 300
            render_dpi = min(max(log_x, 150), target_dpi)

            hdc.StartDoc(job_name)
            doc_started = True

            colorspace = fitz.csGRAY if grayscale else fitz.csRGB
            img_mode = "L" if grayscale else "RGB"

            for _ in range(max(1, copies)):
                for pno in range(len(doc)):
                    page = doc[pno]
                    pix = page.get_pixmap(dpi=render_dpi, colorspace=colorspace, alpha=False)
                    img = Image.frombytes(img_mode, [pix.width, pix.height], pix.samples)
                    dib = ImageWin.Dib(img)

                    hdc.StartPage()
                    scale = min(horz_res / pix.width, vert_res / pix.height)
                    dest_w = int(pix.width * scale)
                    dest_h = int(pix.height * scale)
                    off_x = max(0, (horz_res - dest_w) // 2)
                    off_y = max(0, (vert_res - dest_h) // 2)

                    dib.draw(hdc.GetHandleOutput(), (off_x, off_y, off_x + dest_w, off_y + dest_h))
                    hdc.EndPage()
                    del dib, img, pix

            hdc.EndDoc()
            doc_started = False
            return True
        except Exception:
            if doc_started and hdc is not None:
                try:
                    hdc.AbortDoc()
                except Exception:
                    pass
            return False
        finally:
            doc.close()
            if hdc is not None:
                try:
                    hdc.DeleteDC()
                except Exception:
                    pass

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

        is_virtual = any(v in target.lower() for v in ["virtual", "buffer", "pdf", "xps", "onenote", "fax"])
        if is_virtual:
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

            return {
                "printer": target,
                "supported_bins": bins,
                "supports_duplex": supports_duplex,
                "is_virtual": False,
            }
        except Exception:
            return {
                "printer": target,
                "supported_bins": (),
                "supports_duplex": False,
                "is_virtual": True,
            }

    @classmethod
    def _validate_devmode_options(
        cls,
        printer_name: str,
        tray_id: Optional[int] = None,
        duplex: Optional[str] = None,
    ) -> Optional[Tuple[Optional[int], Optional[int]]]:
        """Validate tray and duplex settings against Guard G2 device capabilities."""
        if sys.platform != "win32" or win32print is None or win32con is None:
            return None

        caps = cls.get_device_capabilities(printer_name)

        valid_tray: Optional[int] = None
        if tray_id is not None:
            if tray_id not in caps.get("supported_bins", []):
                return None
            valid_tray = tray_id

        valid_duplex: Optional[int] = None
        if duplex:
            if not caps.get("supports_duplex", False):
                return None
            dup_lower = duplex.lower()
            if dup_lower in ("vertical", "long", "long-edge", "2-sided-long", "duplex_long"):
                valid_duplex = win32con.DMDUP_VERTICAL
            elif dup_lower in ("horizontal", "short", "short-edge", "2-sided-short", "duplex_short"):
                valid_duplex = win32con.DMDUP_HORIZONTAL
            elif dup_lower in ("simplex", "none", "single", "1-sided", "1-page"):
                valid_duplex = win32con.DMDUP_SIMPLEX
            else:
                return None

        if valid_tray is None and valid_duplex is None:
            return None
        return valid_tray, valid_duplex

    @classmethod
    def build_devmode(
        cls,
        printer_name: str,
        tray_id: Optional[int] = None,
        duplex: Optional[str] = None,
    ) -> Optional[Any]:
        """
        Build a per-job local PyDEVMODE via DocumentProperties without requiring
        PRINTER_ALL_ACCESS or mutating system-wide default printer settings via SetPrinter.
        """
        validated = cls._validate_devmode_options(printer_name, tray_id=tray_id, duplex=duplex)
        if validated is None:
            return None
        valid_tray, valid_duplex = validated

        try:
            h_printer = win32print.OpenPrinter(printer_name)
            try:
                p_info = win32print.GetPrinter(h_printer, 2)
                devmode = p_info.get("pDevMode")
                if devmode is None:
                    return None
                if valid_tray is not None:
                    devmode.DefaultSource = valid_tray
                    devmode.Fields |= getattr(win32con, "DM_DEFAULTSOURCE", 0x00000200)
                if valid_duplex is not None:
                    devmode.Duplex = valid_duplex
                    devmode.Fields |= getattr(win32con, "DM_DUPLEX", 0x00001000)

                dm_in_out = getattr(win32con, "DM_IN_BUFFER", 8) | getattr(win32con, "DM_OUT_BUFFER", 2)
                win32print.DocumentProperties(0, h_printer, printer_name, devmode, devmode, dm_in_out)
                return devmode
            finally:
                win32print.ClosePrinter(h_printer)
        except Exception:
            return None

    @classmethod
    def apply_devmode(
        cls,
        printer_name: str,
        tray_id: Optional[int] = None,
        duplex: Optional[str] = None,
    ) -> bool:
        """
        Safely validates and prepares local printer DEVMODE with Guard G2 capabilities check
        (non-privileged, per-job safe without global SetPrinter mutation).
        """
        return cls.build_devmode(printer_name, tray_id=tray_id, duplex=duplex) is not None

    @classmethod
    def get_printer_status(cls, printer_name: Optional[str] = None) -> Dict[str, Any]:
        """
        Queries printer hardware status, paper availability, and spooler job queue.
        Detects JOB_STATUS_PAPEROUT, PRINTER_STATUS_PAPER_JAM, OFFLINE, and DOOR_OPEN.
        """
        target = printer_name or cls.get_default_printer() or "Virtual Spooler (PDF Buffer)"
        is_virtual = any(v in target.lower() for v in ["virtual", "buffer", "pdf", "xps", "onenote", "fax"])

        if is_virtual or sys.platform != "win32" or win32print is None:
            return {
                "printer": target,
                "status": "READY",
                "message": "Máy in sẵn sàng tiếp nhận lệnh in",
                "is_virtual": is_virtual,
                "c_jobs": 0,
                "jobs": [],
                "hardware_warning": None,
                "supports_duplex": False if is_virtual else True,
            }

        try:
            h_printer = win32print.OpenPrinter(target)
            try:
                p_info = win32print.GetPrinter(h_printer, 2)
                raw_status = p_info.get("Status", 0)
                c_jobs = p_info.get("cJobs", 0)

                jobs_list: List[Dict[str, Any]] = []
                try:
                    raw_jobs = win32print.EnumJobs(h_printer, 0, 10, 2)
                    for j in raw_jobs:
                        jobs_list.append({
                            "job_id": j.get("JobId"),
                            "document": j.get("pDocument", "Document"),
                            "status": j.get("Status", 0),
                            "total_pages": j.get("TotalPages", 0),
                            "pages_printed": j.get("PagesPrinted", 0),
                        })
                except Exception:
                    pass

                caps = cls.get_device_capabilities(target)

                warning_msg: Optional[str] = None
                status_code = "READY"
                status_desc = "Máy in sẵn sàng tiếp nhận lệnh in"

                paper_out_val = getattr(win32print, "PRINTER_STATUS_PAPER_OUT", 0x00000010)
                paper_jam_val = getattr(win32print, "PRINTER_STATUS_PAPER_JAM", 0x00000008)
                offline_val = getattr(win32print, "PRINTER_STATUS_OFFLINE", 0x00000080)
                door_open_val = getattr(win32print, "PRINTER_STATUS_DOOR_OPEN", 0x00400000)
                error_val = getattr(win32print, "PRINTER_STATUS_ERROR", 0x00000002)
                paused_val = getattr(win32print, "PRINTER_STATUS_PAUSED", 0x00000001)

                if raw_status & paper_out_val:
                    status_code = "PAPER_OUT"
                    status_desc = "Máy in hết giấy (JOB_STATUS_PAPEROUT)"
                    warning_msg = "Khay giấy của máy in đang rỗng. Vui lòng nạp thêm giấy vào khay."
                elif raw_status & paper_jam_val:
                    status_code = "PAPER_JAM"
                    status_desc = "Máy in bị kẹt giấy (JOB_STATUS_PAPER_JAM)"
                    warning_msg = "Phát hiện kẹt giấy trong cơ cấu cuốn giấy của máy in."
                elif raw_status & offline_val:
                    status_code = "OFFLINE"
                    status_desc = "Máy in đang ngoại tuyến (Printer Offline)"
                    warning_msg = "Máy in chưa kết nối hoặc chưa bật nguồn."
                elif raw_status & door_open_val:
                    status_code = "DOOR_OPEN"
                    status_desc = "Nắp máy in đang mở (Door Open)"
                    warning_msg = "Nắp khoang mực hoặc khay đảo mặt đang mở."
                elif raw_status & error_val:
                    status_code = "ERROR"
                    status_desc = "Máy in đang báo lỗi phần cứng"
                    warning_msg = "Máy in gặp sự cố kỹ thuật hoặc lỗi driver."
                elif raw_status & paused_val:
                    status_code = "PAUSED"
                    status_desc = "Hàng đợi in đang tạm dừng (Spooler Paused)"
                elif c_jobs > 0:
                    status_desc = f"Máy in đang xử lý {c_jobs} lệnh in trong hàng đợi"

                return {
                    "printer": target,
                    "status": status_code,
                    "message": status_desc,
                    "is_virtual": False,
                    "c_jobs": c_jobs,
                    "jobs": jobs_list,
                    "hardware_warning": warning_msg,
                    "supports_duplex": caps.get("supports_duplex", False),
                    "supported_bins": caps.get("supported_bins", ()),
                }
            finally:
                win32print.ClosePrinter(h_printer)
        except Exception as e:
            return {
                "printer": target,
                "status": "UNKNOWN",
                "message": f"Không thể kết nối máy in: {e}",
                "is_virtual": False,
                "c_jobs": 0,
                "jobs": [],
                "hardware_warning": f"Lỗi đọc trạng thái spooler: {e}",
                "supports_duplex": False,
            }

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
        page_range: Optional[Union[str, List[int]]] = None,
        page_size: Optional[str] = None,
        scaling: Optional[str] = "fit",
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

        # Pre-check printer status for hardware warnings
        pre_status = cls.get_printer_status(target_printer)
        hardware_warning = pre_status.get("hardware_warning")

        # Step 0: Build per-job local DEVMODE settings if physical printer on Windows (Guard G2)
        local_devmode = None
        if sys.platform == "win32" and (tray_id is not None or duplex is not None):
            local_devmode = cls.build_devmode(target_printer, tray_id=tray_id, duplex=duplex)
        devmode_applied = local_devmode is not None

        temp_files_to_clean: List[str] = []
        async_spool = False

        try:
            # Step 1: Ensure PDF format for printing
            ext = os.path.splitext(file_path)[1].lower()
            printable_pdf = file_path

            if ext != ".pdf":
                temp_pdf = os.path.join(tempfile.gettempdir(), f"{job_id}_converted.pdf")
                temp_files_to_clean.append(temp_pdf)
                converted_ok = False
                if ext in (".docx", ".pptx", ".xlsx", ".xls", ".xlsm"):
                    try:
                        from core.converter.pdf_renderer import UniversalPdfRenderer
                        rendered = UniversalPdfRenderer().convert_to_pdf(file_path, temp_pdf)
                        if os.path.isfile(rendered) and os.path.getsize(rendered) > 0:
                            printable_pdf = rendered
                            converted_ok = True
                    except Exception:
                        converted_ok = False

                if not converted_ok:
                    try:
                        from core.converter.engine import convert_universal
                        ok_conv, out_conv, _elap = convert_universal(file_path, temp_pdf)
                        if not ok_conv or not os.path.isfile(out_conv):
                            raise RuntimeError(out_conv)
                        printable_pdf = out_conv
                    except Exception as ce:
                        raise RuntimeError(f"Không thể chuyển đổi {ext} sang PDF để in: {ce}")

            # Step 2: Apply Print Optimizer (Vector Grayscale, N-Up, Booklet Creep, Watermark, Margin Guard, Page Range, Page Size)
            opt_dict = dict(optimizer_options or {})
            if page_range:
                opt_dict["page_range"] = page_range
            if page_size:
                opt_dict["page_size"] = page_size
            if scaling:
                opt_dict["scaling"] = scaling

            opt_stats: Dict[str, Any] = {}
            if opt_dict:
                optimized_pdf = os.path.join(tempfile.gettempdir(), f"{job_id}_optimized.pdf")
                temp_files_to_clean.append(optimized_pdf)
                opt_stats = PrintOptimizer.optimize_pdf(
                    input_pdf_path=printable_pdf,
                    output_pdf_path=optimized_pdf,
                    grayscale=opt_dict.get("grayscale", False),
                    vector_grayscale=opt_dict.get("vector_grayscale", True),
                    dpi=opt_dict.get("dpi", 300),
                    n_up=opt_dict.get("n_up", 1),
                    booklet=opt_dict.get("booklet", False),
                    creep_compensation=opt_dict.get("creep_compensation", True),
                    paper_thickness_pt=opt_dict.get("paper_thickness_pt", 0.28),
                    watermark_text=opt_dict.get("watermark"),
                    margin_guard_mm=opt_dict.get("margin_guard_mm", 5.0),
                    page_range=opt_dict.get("page_range"),
                    page_size=opt_dict.get("page_size"),
                    scaling=opt_dict.get("scaling", "fit"),
                )
                printable_pdf = optimized_pdf

            # Base response payload
            base_res: Dict[str, Any] = {
                "success": True,
                "job_id": job_id,
                "printer": target_printer,
                "copies": copies,
                "hardware_warning": hardware_warning,
                "printer_status": pre_status.get("status", "READY"),
                "devmode_applied": devmode_applied,
                "duplex": duplex or "simplex",
                "page_range": opt_dict.get("page_range", "all"),
                "page_size": opt_dict.get("page_size", "A4"),
                "scaling": opt_dict.get("scaling", "fit"),
                "optimization": opt_stats,
            }

            # Step 3: Check if Virtual Spooler is requested or if printer is Virtual Spooler
            if "virtual" in target_printer.lower() or "buffer" in target_printer.lower():
                spool_dir = os.path.abspath(os.path.join(os.getcwd(), "output", "spool"))
                os.makedirs(spool_dir, exist_ok=True)
                spool_file = os.path.join(spool_dir, f"{job_id}_{os.path.basename(file_path)}.pdf")
                shutil.copy2(printable_pdf, spool_file)
                return {
                    **base_res,
                    "spooled_file": spool_file,
                    "mode": "VIRTUAL_SPOOL",
                    "status": "QUEUED"
                }

            # Step 4: Physical Spooling on Windows
            if sys.platform == "win32":
                # Priority 1: Native Windows GDI Spooler (per-job DEVMODE + adaptive 600 DPI / csGRAY)
                try:
                    gdi_ok = cls._spool_via_gdi(
                        printable_pdf,
                        target_printer,
                        job_name=job_name,
                        copies=copies,
                        devmode=local_devmode,
                        grayscale=bool(opt_dict.get("grayscale", False)),
                    )
                    if gdi_ok:
                        return {
                            **base_res,
                            "mode": "WIN32_GDI_DIRECT",
                            "status": "SENT_TO_PRINTER"
                        }
                except Exception:
                    pass

                # Priority 2: ShellExecute printto (async handoff, keep temp pdf until OS picks it up)
                if win32api is not None:
                    try:
                        for _ in range(copies):
                            win32api.ShellExecute(0, "printto", printable_pdf, f'"{target_printer}"', ".", 0)
                        async_spool = True
                        return {
                            **base_res,
                            "mode": "SHELLEXECUTE_PRINTTO",
                            "status": "SENT_TO_PRINTER"
                        }
                    except Exception:
                        pass

                # Priority 3: Headless Edge (with safe timeout and exception catch)
                edge_bin = cls._find_edge_binary()
                if edge_bin and silent:
                    try:
                        cmd = [
                            edge_bin,
                            "--headless",
                            "--disable-gpu",
                            f"--print-to-printer={target_printer}",
                            printable_pdf
                        ]
                        for _ in range(copies):
                            subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=10)
                        return {
                            **base_res,
                            "mode": "EDGE_HEADLESS_SPOOL",
                            "status": "SENT_TO_PRINTER"
                        }
                    except Exception:
                        pass

                # Priority 4: Graceful Virtual Spool fallback if hardware driver blocks execution
                spool_dir = os.path.abspath(os.path.join(os.getcwd(), "output", "spool"))
                os.makedirs(spool_dir, exist_ok=True)
                spool_file = os.path.join(spool_dir, f"{job_id}_{os.path.basename(file_path)}.pdf")
                shutil.copy2(printable_pdf, spool_file)
                return {
                    **base_res,
                    "spooled_file": spool_file,
                    "mode": "VIRTUAL_SPOOL_FALLBACK",
                    "status": "QUEUED",
                    "hardware_warning": f"Không thể gửi lệnh in trực tiếp đến driver '{target_printer}'. Tệp in đã được đệm an toàn tại output/spool/."
                }

            # Step 5: Physical Spooling on Linux / CUPS
            else:
                cmd = ["lp", "-d", target_printer, "-n", str(copies), "-t", job_name, printable_pdf]
                res = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, timeout=15)
                if res.returncode != 0:
                    raise RuntimeError(f"Lỗi gửi lệnh in CUPS (lp): {res.stderr}")
                return {
                    **base_res,
                    "mode": "CUPS_LP",
                    "status": "SENT_TO_PRINTER"
                }
        finally:
            if not async_spool:
                for tmp_p in temp_files_to_clean:
                    if tmp_p and os.path.isfile(tmp_p):
                        try:
                            os.remove(tmp_p)
                        except OSError:
                            pass
