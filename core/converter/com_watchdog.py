# -*- coding: utf-8 -*-
"""
core.converter.com_watchdog — Win32 COM Process Isolation & Timeout Watchdog.

Conforms to Antigravity v5.0 Enterprise Architecture:
- Thread-isolated execution with strict timeout (default 15s).
- Automatic COM threading model initialization (CoInitialize / CoUninitialize).
- Active unfreeze & orphaned process reaper (taskkill /F /IM WINWORD.EXE /IM POWERPNT.EXE).
- Graceful exception translation for fallback to LibreOffice Headless.
"""

from __future__ import annotations

import os
import subprocess
import sys
import threading
from typing import Any, Callable, Dict, Optional, Tuple


class ComProcessWatchdog:
    """Manages thread-isolated COM automation calls with watchdog timeout protection."""

    PROCESS_MAP: Dict[str, str] = {
        "word": "WINWORD.EXE",
        "powerpoint": "POWERPNT.EXE",
        "excel": "EXCEL.EXE",
    }

    @classmethod
    def terminate_hung_com_processes(cls, app_type: str = "all") -> int:
        """
        Force-terminates hung Microsoft Office processes on Windows.
        Returns the number of target process types targeted.
        """
        if sys.platform != "win32":
            return 0

        targets = []
        app_lower = app_type.lower()
        if "word" in app_lower:
            targets.append("WINWORD.EXE")
        elif "powerpoint" in app_lower or "pptx" in app_lower:
            targets.append("POWERPNT.EXE")
        elif "excel" in app_lower or "xlsx" in app_lower:
            targets.append("EXCEL.EXE")
        else:
            targets = ["WINWORD.EXE", "POWERPNT.EXE", "EXCEL.EXE"]

        killed = 0
        for proc in targets:
            try:
                cmd = ["taskkill", "/F", "/IM", proc]
                subprocess.run(
                    cmd,
                    stdout=subprocess.DEVNULL,
                    stderr=subprocess.DEVNULL,
                    timeout=5,
                    check=False,
                )
                killed += 1
            except Exception:
                pass
        return killed

    @classmethod
    def run_with_timeout(
        cls,
        func: Callable[..., Any],
        args: Tuple[Any, ...] = (),
        kwargs: Optional[Dict[str, Any]] = None,
        timeout_sec: float = 15.0,
        app_type: str = "word",
    ) -> Any:
        """
        Executes a COM automation callable in an isolated thread with CoInitialize
        and force-kills the process if it hangs beyond timeout_sec.
        """
        if kwargs is None:
            kwargs = {}

        result_container: Dict[str, Any] = {}
        exception_container: Dict[str, Exception] = {}

        def _worker() -> None:
            # Multi-threaded COM initialization
            if sys.platform == "win32":
                try:
                    import pythoncom  # type: ignore
                    pythoncom.CoInitialize()
                except Exception:
                    pass

            try:
                result_container["result"] = func(*args, **kwargs)
            except Exception as e:
                exception_container["error"] = e
            finally:
                if sys.platform == "win32":
                    try:
                        import pythoncom  # type: ignore
                        pythoncom.CoUninitialize()
                    except Exception:
                        pass

        thread = threading.Thread(target=_worker, daemon=True)
        thread.start()
        thread.join(timeout=timeout_sec)

        if thread.is_alive():
            # Process is hung (modal dialog, activation prompt, deadlock)
            cls.terminate_hung_com_processes(app_type)
            raise TimeoutError(
                f"COM automation for '{app_type}' hung and timed out after {timeout_sec}s. "
                f"Orphaned processes terminated to prevent system unresponsiveness."
            )

        if "error" in exception_container:
            raise exception_container["error"]

        return result_container.get("result")

