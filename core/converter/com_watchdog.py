# -*- coding: utf-8 -*-
"""
core.converter.com_watchdog — Win32 COM Process Isolation & Timeout Watchdog.

Conforms to Antigravity v5.0 Enterprise Architecture:
- Thread-isolated execution with strict timeout (default 15s).
- Automatic COM threading model initialization (CoInitialize / CoUninitialize).
- PID-scoped hung COM reaper (terminates only registered worker PIDs or DCOM /automation -Embedding
  processes spawned during the hung call, never touching user-opened interactive Office windows).
- Graceful exception translation for fallback to LibreOffice Headless.
"""

from __future__ import annotations

import json
import os
import subprocess
import sys
import threading
from typing import Any, Callable, Dict, Iterable, List, Optional, Set, Tuple


class ComProcessWatchdog:
    """Manages thread-isolated COM automation calls with watchdog timeout protection."""

    PROCESS_MAP: Dict[str, str] = {
        "word": "WINWORD.EXE",
        "powerpoint": "POWERPNT.EXE",
        "excel": "EXCEL.EXE",
    }
    AUTOMATION_FLAGS: Tuple[str, ...] = ("-embedding", "/automation")

    _TRACKING_LOCK = threading.Lock()
    _ACTIVE_WORKER_PIDS: Dict[int, Set[int]] = {}

    @classmethod
    def register_current_pid(cls, pid: int) -> None:
        """Associate a spawned COM server PID with the current watchdog worker thread."""
        if pid <= 0:
            return
        tid = threading.get_ident()
        with cls._TRACKING_LOCK:
            if tid in cls._ACTIVE_WORKER_PIDS:
                cls._ACTIVE_WORKER_PIDS[tid].add(int(pid))

    @classmethod
    def register_pid_from_hwnd(cls, hwnd: Any) -> Optional[int]:
        """Resolve PID from a COM Application window handle (HWND) and register it for cleanup."""
        if sys.platform != "win32" or not hwnd:
            return None
        try:
            import ctypes
            pid_c = ctypes.c_ulong(0)
            ctypes.windll.user32.GetWindowThreadProcessId(int(hwnd), ctypes.byref(pid_c))
            pid = int(pid_c.value)
            if pid > 0:
                cls.register_current_pid(pid)
                return pid
        except Exception:
            pass
        return None

    @classmethod
    def _resolve_target_process_names(cls, app_type: str = "all") -> List[str]:
        app_lower = (app_type or "all").lower()
        if "word" in app_lower:
            return ["WINWORD.EXE"]
        if "powerpoint" in app_lower or "pptx" in app_lower:
            return ["POWERPNT.EXE"]
        if "excel" in app_lower or "xlsx" in app_lower or "xls" in app_lower:
            return ["EXCEL.EXE"]
        return ["WINWORD.EXE", "POWERPNT.EXE", "EXCEL.EXE"]

    @classmethod
    def _is_automation_cmdline(cls, cmdline: str) -> bool:
        low = (cmdline or "").lower()
        return any(flag in low for flag in cls.AUTOMATION_FLAGS)

    @classmethod
    def _list_office_processes(
        cls, process_names: Optional[List[str]] = None
    ) -> List[Tuple[int, str, str]]:
        """Query Windows processes filtered to Office executables with their command lines."""
        if sys.platform != "win32":
            return []
        targets = {p.lower() for p in (process_names or cls.PROCESS_MAP.values())}
        ps_cmd = (
            "Get-CimInstance Win32_Process | "
            "Where-Object { $_.Name -match '^(WINWORD|EXCEL|POWERPNT)\\.EXE$' } | "
            "Select-Object ProcessId,Name,CommandLine | ConvertTo-Json -Compress"
        )
        try:
            res = subprocess.run(
                ["powershell", "-NoProfile", "-NonInteractive", "-Command", ps_cmd],
                capture_output=True,
                text=True,
                timeout=8,
                errors="replace",
            )
            if res.returncode != 0 or not res.stdout.strip():
                return []
            raw = json.loads(res.stdout)
            items = raw if isinstance(raw, list) else [raw]
            parsed: List[Tuple[int, str, str]] = []
            for item in items:
                name = str(item.get("Name") or "")
                if name.lower() not in targets:
                    continue
                pid = int(item.get("ProcessId") or 0)
                cmdline = str(item.get("CommandLine") or "")
                if pid > 0:
                    parsed.append((pid, name, cmdline))
            return parsed
        except Exception:
            return []

    @classmethod
    def _kill_pid(cls, pid: int) -> bool:
        """Terminate a single Windows process by numeric PID."""
        try:
            proc = subprocess.run(
                ["taskkill", "/F", "/PID", str(int(pid))],
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
                timeout=5,
                check=False,
            )
            return proc.returncode == 0
        except Exception:
            return False

    @classmethod
    def terminate_hung_com_processes(
        cls,
        app_type: str = "all",
        pids: Optional[Iterable[int]] = None,
        automation_only: bool = True,
    ) -> int:
        """
        Safely terminates hung Microsoft Office COM automation processes on Windows.
        - If explicit `pids` are supplied, terminates only those exact PIDs.
        - Otherwise when `automation_only=True` (default), terminates only COM-launched instances
          carrying `-Embedding` or `/automation`, leaving interactive user documents untouched.
        Returns the number of terminated processes/targets.
        """
        if sys.platform != "win32":
            return 0

        if pids is not None:
            killed = 0
            for pid in sorted({int(p) for p in pids if int(p) > 0}):
                if cls._kill_pid(pid):
                    killed += 1
            return killed

        targets = cls._resolve_target_process_names(app_type)
        if automation_only:
            killed = 0
            for pid, _name, cmdline in cls._list_office_processes(targets):
                if cls._is_automation_cmdline(cmdline):
                    if cls._kill_pid(pid):
                        killed += 1
            return killed

        killed = 0
        for proc in targets:
            try:
                subprocess.run(
                    ["taskkill", "/F", "/IM", proc],
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
        and force-kills only its tracked/spawned COM PID(s) if it hangs beyond timeout_sec.
        """
        if kwargs is None:
            kwargs = {}

        target_names = cls._resolve_target_process_names(app_type)
        pre_pids = {pid for pid, _, _ in cls._list_office_processes(target_names)}

        result_container: Dict[str, Any] = {}
        exception_container: Dict[str, Exception] = {}
        worker_tid_box: List[int] = []

        def _worker() -> None:
            tid = threading.get_ident()
            worker_tid_box.append(tid)
            with cls._TRACKING_LOCK:
                cls._ACTIVE_WORKER_PIDS[tid] = set()

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
                with cls._TRACKING_LOCK:
                    # Keep tracked PIDs briefly if parent timed out? Parent reads before clearing when alive,
                    # or clears after join handles timeout. Safe to discard only when exiting normally without timeout.
                    pass
                if sys.platform == "win32":
                    try:
                        import pythoncom  # type: ignore
                        pythoncom.CoUninitialize()
                    except Exception:
                        pass

        thread = threading.Thread(target=_worker, daemon=True)
        thread.start()
        thread.join(timeout=timeout_sec)

        tid = worker_tid_box[0] if worker_tid_box else None
        if thread.is_alive():
            with cls._TRACKING_LOCK:
                tracked = set(cls._ACTIVE_WORKER_PIDS.pop(tid, set())) if tid is not None else set()
            post_auto = {
                pid
                for pid, _, cmdline in cls._list_office_processes(target_names)
                if pid not in pre_pids and cls._is_automation_cmdline(cmdline)
            }
            kill_targets = tracked | post_auto
            if kill_targets:
                cls.terminate_hung_com_processes(app_type, pids=kill_targets)
            raise TimeoutError(
                f"COM automation for '{app_type}' hung and timed out after {timeout_sec}s. "
                f"Orphaned COM processes terminated to prevent system unresponsiveness."
            )

        if tid is not None:
            with cls._TRACKING_LOCK:
                cls._ACTIVE_WORKER_PIDS.pop(tid, None)

        if "error" in exception_container:
            raise exception_container["error"]

        return result_container.get("result")

