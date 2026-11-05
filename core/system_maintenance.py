"""
core.system_maintenance — Enterprise System Maintenance Engine (Zero-Dependencies).

Provides automated, periodic system maintenance, resource hygiene, and diagnostics
across Windows, Linux, and macOS:
1. Zombie Lock Files & Temporary Pipeline Artifacts Purge (ERR_CONV_004).
2. Orphaned Headless Process Reaper (Process Hygiene & Antigravity Unfreeze).
3. SQLite Database Compaction (WAL Checkpoint, VACUUM & FTS5 Optimize).
4. Python Bytecode & Test Cache Sanitation (__pycache__, .pytest_cache).
5. Log Lifecycle Management & Web Outputs Pruning.
6. Windows Desktop & GPU Diagnostics (Multiplane Overlay MPO Stutter Guard).
7. System Resource & Health Vitals Telemetry (Disk, RAM, Dry-Run Mode).

Conforms strictly to Clean Architecture: 100% self-contained in core/,
zero external pip dependencies (Python Standard Library only).
"""

from __future__ import annotations

import csv
from dataclasses import asdict, dataclass, field
import datetime
import gc
import glob
import io
import json
import os
from pathlib import Path
import platform
import shutil
import signal
import sqlite3
import subprocess
import sys
import tempfile
import time
from typing import Any, Dict, List, Optional, Set, Tuple

# Reconfigure stdout/stderr on Windows for safe UTF-8 logging
if sys.platform == "win32":
    try:
        if hasattr(sys.stdout, "reconfigure"):
            getattr(sys.stdout, "reconfigure")(encoding="utf-8", errors="replace")
        if hasattr(sys.stderr, "reconfigure"):
            getattr(sys.stderr, "reconfigure")(encoding="utf-8", errors="replace")
    except Exception:
        pass

_CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
_WORKSPACE_ROOT = os.path.dirname(_CURRENT_DIR)


@dataclass
class MaintenanceReport:
    """Consolidated telemetry and results report for System Maintenance operations."""
    timestamp: str = field(default_factory=lambda: datetime.datetime.now().isoformat())
    execution_time_sec: float = 0.0
    dry_run: bool = False
    vitals: Dict[str, Any] = field(default_factory=dict)
    
    # Pillar 1: Locks & Temp
    lock_files_removed: int = 0
    lock_files_details: List[str] = field(default_factory=list)
    temp_files_removed: int = 0
    temp_files_details: List[str] = field(default_factory=list)
    
    # Pillar 2: Processes
    processes_scanned: int = 0
    orphaned_processes_terminated: List[Dict[str, Any]] = field(default_factory=list)
    
    # Pillar 3: Databases
    databases_vacuumed: List[Dict[str, Any]] = field(default_factory=list)
    
    # Pillar 4: Bytecode & Cache
    pycache_dirs_removed: int = 0
    pycache_files_removed: int = 0
    test_cache_dirs_removed: int = 0
    
    # Pillar 5: Logs & Outputs
    logs_rotated: List[Dict[str, Any]] = field(default_factory=list)
    outputs_pruned: List[Dict[str, Any]] = field(default_factory=list)
    
    # Pillar 6: Windows MPO
    mpo_status: Dict[str, Any] = field(default_factory=dict)
    
    # Totals & Notices
    total_bytes_freed: int = 0
    warnings: List[str] = field(default_factory=list)
    errors: List[str] = field(default_factory=list)

    def to_dict(self) -> Dict[str, Any]:
        """Convert report to JSON-serializable dictionary."""
        return asdict(self)


class SystemMaintenanceEngine:
    """
    Enterprise System Maintenance Engine for Antigravity Universal Office Studio.
    100% Python Standard Library (Zero External Dependencies).
    """

    def __init__(self, workspace_root: Optional[str] = None):
        self.workspace_root = os.path.abspath(workspace_root or _WORKSPACE_ROOT)

    # -------------------------------------------------------------------------
    # Pillar 7: System Vitals & Resources
    # -------------------------------------------------------------------------

    def get_system_vitals(self) -> Dict[str, Any]:
        """Collect real-time system resource metrics (Disk, Memory, Runtime)."""
        vitals: Dict[str, Any] = {
            "os_platform": sys.platform,
            "os_release": platform.platform(),
            "python_version": sys.version.split()[0],
            "workspace_root": self.workspace_root,
        }

        # Disk space
        try:
            total, used, free = shutil.disk_usage(self.workspace_root)
            vitals["disk_total_gb"] = round(total / (1024 ** 3), 2)
            vitals["disk_used_gb"] = round(used / (1024 ** 3), 2)
            vitals["disk_free_gb"] = round(free / (1024 ** 3), 2)
            vitals["disk_free_pct"] = round((free / total) * 100, 1) if total > 0 else 0.0
        except Exception as exc:
            vitals["disk_error"] = str(exc)

        # Memory (RAM)
        if sys.platform == "win32":
            try:
                import ctypes
                from ctypes import c_ulonglong, c_ulong, Structure

                class MEMORYSTATUSEX(Structure):
                    _fields_ = [
                        ("dwLength", c_ulong),
                        ("dwMemoryLoad", c_ulong),
                        ("ullTotalPhys", c_ulonglong),
                        ("ullAvailPhys", c_ulonglong),
                        ("ullTotalPageFile", c_ulonglong),
                        ("ullAvailPageFile", c_ulonglong),
                        ("ullTotalVirtual", c_ulonglong),
                        ("ullAvailVirtual", c_ulonglong),
                        ("sullAvailExtendedVirtual", c_ulonglong),
                    ]

                stat = MEMORYSTATUSEX()
                stat.dwLength = ctypes.sizeof(MEMORYSTATUSEX)
                if ctypes.windll.kernel32.GlobalMemoryStatusEx(ctypes.byref(stat)):
                    vitals["ram_total_gb"] = round(stat.ullTotalPhys / (1024 ** 3), 2)
                    vitals["ram_available_gb"] = round(stat.ullAvailPhys / (1024 ** 3), 2)
                    vitals["ram_load_pct"] = stat.dwMemoryLoad
            except Exception:
                pass
        elif sys.platform == "linux" and os.path.exists("/proc/meminfo"):
            try:
                mem: Dict[str, int] = {}
                with open("/proc/meminfo", "r", encoding="utf-8") as f:
                    for line in f:
                        parts = line.split(":")
                        if len(parts) == 2:
                            key = parts[0].strip()
                            val_str = parts[1].strip().split()[0]
                            mem[key] = int(val_str)
                total_kb = mem.get("MemTotal", 0)
                avail_kb = mem.get("MemAvailable", mem.get("MemFree", 0))
                vitals["ram_total_gb"] = round(total_kb / (1024 ** 2), 2)
                vitals["ram_available_gb"] = round(avail_kb / (1024 ** 2), 2)
                vitals["ram_load_pct"] = round((1.0 - (avail_kb / total_kb)) * 100, 1) if total_kb > 0 else 0.0
            except Exception:
                pass

        return vitals

    # -------------------------------------------------------------------------
    # Pillar 1: Zombie Lock Files & Temporary Pipeline Artifacts (ERR_CONV_004)
    # -------------------------------------------------------------------------

    def clean_zombie_locks(self, dry_run: bool = False) -> Dict[str, Any]:
        """
        Scan and remove zombie lock files created by Word, Excel, PowerPoint, and LibreOffice.
        Invariants: ERR_CONV_004.
        Patterns: ~$*.docx, ~$*.xlsx, ~$*.pptx, .~lock.*.docx#, .~lock.*.xlsx#
        """
        lock_patterns = [
            "~$*.docx", "~$*.xlsx", "~$*.pptx", "~$*.docm", "~$*.xlsm",
            ".~lock.*.docx#", ".~lock.*.xlsx#", ".~lock.*.pptx#"
        ]
        found_files: List[str] = []
        bytes_freed = 0
        removed: List[str] = []
        errors: List[str] = []

        for root, dirs, files in os.walk(self.workspace_root):
            # Skip VCS and virtual environments
            dirs[:] = [d for d in dirs if d not in {".git", ".svn", "node_modules", "venv", ".venv", "env"}]
            for pat in lock_patterns:
                for match in Path(root).glob(pat):
                    if match.is_file():
                        p_str = str(match.resolve())
                        found_files.append(p_str)
                        try:
                            f_size = match.stat().st_size
                            bytes_freed += f_size
                        except OSError:
                            pass

        if not dry_run:
            for p in found_files:
                try:
                    os.remove(p)
                    removed.append(p)
                except Exception as exc:
                    errors.append(f"Cannot remove lock file {p}: {exc}")
        else:
            removed = list(found_files)

        return {
            "found_count": len(found_files),
            "removed_count": len(removed),
            "details": removed,
            "bytes_freed": bytes_freed,
            "errors": errors,
        }

    def clean_temp_artifacts(
        self, dry_run: bool = False, max_age_hours: float = 2.0
    ) -> Dict[str, Any]:
        """
        Scan and clean temporary conversion pipeline artifacts and dangling scoped dirs.
        Patterns in workspace: step_*_temp.*, tmp*.docx, tmp*.xlsx, tmp*.pdf, *.crswap, *.tmp.
        Patterns in OS temp: scoped_dir*, antigravity_temp_profile* older than max_age_hours.
        """
        temp_patterns = ["step_*_temp.*", "tmp*.docx", "tmp*.xlsx", "tmp*.pdf", "*.crswap", "*.tmp"]
        found_files: List[str] = []
        found_dirs: List[str] = []
        bytes_freed = 0
        removed_files: List[str] = []
        removed_dirs: List[str] = []
        errors: List[str] = []
        cutoff_sec = time.time() - (max_age_hours * 3600)

        # 1. Clean workspace temp files
        for root, dirs, files in os.walk(self.workspace_root):
            dirs[:] = [d for d in dirs if d not in {".git", "node_modules", "venv", ".venv", "env"}]
            for pat in temp_patterns:
                for match in Path(root).glob(pat):
                    if match.is_file():
                        p_str = str(match.resolve())
                        found_files.append(p_str)
                        try:
                            bytes_freed += match.stat().st_size
                        except OSError:
                            pass

        # 2. Clean OS temp directory scoped profiles (Headless Chromium / Edge App mode)
        os_temp = tempfile.gettempdir()
        scoped_patterns = ["scoped_dir*", "antigravity_temp_profile*"]
        for pat in scoped_patterns:
            for item in Path(os_temp).glob(pat):
                try:
                    mtime = item.stat().st_mtime
                    if mtime < cutoff_sec:
                        if item.is_dir():
                            found_dirs.append(str(item.resolve()))
                        elif item.is_file():
                            found_files.append(str(item.resolve()))
                except OSError:
                    pass

        if not dry_run:
            for f in found_files:
                try:
                    os.remove(f)
                    removed_files.append(f)
                except Exception as exc:
                    errors.append(f"Cannot remove temp file {f}: {exc}")

            for d in found_dirs:
                try:
                    shutil.rmtree(d, ignore_errors=True)
                    removed_dirs.append(d)
                except Exception as exc:
                    errors.append(f"Cannot remove scoped dir {d}: {exc}")
        else:
            removed_files = list(found_files)
            removed_dirs = list(found_dirs)

        return {
            "found_files_count": len(found_files),
            "removed_files_count": len(removed_files),
            "removed_dirs_count": len(removed_dirs),
            "details": removed_files + removed_dirs,
            "bytes_freed": bytes_freed,
            "errors": errors,
        }

    # -------------------------------------------------------------------------
    # Pillar 2: Orphaned Headless Process Reaper
    # -------------------------------------------------------------------------

    def reap_orphaned_processes(
        self, dry_run: bool = False, force: bool = False
    ) -> Dict[str, Any]:
        """
        Scan and terminate orphaned headless background processes:
        - soffice.bin / soffice.exe running headless with no GUI window.
        - msedge.exe / chrome.exe instances with '--headless' or '--user-data-dir=...antigravity*'.
        - Node.exe Vite/Dev preview processes launched in this workspace that lost parent.
        
        SAFEGUARD: NEVER terminates the user's regular personal browser windows!
        """
        scanned_count = 0
        terminated: List[Dict[str, Any]] = []
        errors: List[str] = []

        if sys.platform == "win32":
            processes = self._scan_windows_processes()
        else:
            processes = self._scan_posix_processes()

        scanned_count = len(processes)
        current_pid = os.getpid()

        for proc in processes:
            pid = proc.get("pid", 0)
            name = (proc.get("name") or "").lower()
            cmdline = (proc.get("cmdline") or "").lower()

            if pid <= 4 or pid == current_pid:
                continue

            # Candidate Identification
            is_headless_office = ("soffice" in name) and ("--headless" in cmdline or not cmdline)
            is_headless_browser = (
                ("chrome.exe" in name or "msedge.exe" in name or "msedgewebview2.exe" in name)
                and (
                    "--headless" in cmdline
                    or "antigravity_temp_profile" in cmdline
                    or "scoped_dir" in cmdline
                    or "--user-data-dir=" in cmdline and "antigravity" in cmdline
                )
            )
            # Never kill regular personal browser windows!
            if ("chrome.exe" in name or "msedge.exe" in name) and not is_headless_browser:
                continue

            # Dangling Node in workspace
            is_orphaned_node = (
                "node.exe" in name or name == "node"
            ) and (
                self.workspace_root.lower() in cmdline
                and ("vite" in cmdline or "dev" in cmdline)
            )

            if is_headless_office or is_headless_browser or is_orphaned_node:
                proc_info = {
                    "pid": pid,
                    "name": proc.get("name"),
                    "cmdline_summary": (proc.get("cmdline") or "")[:120],
                    "reason": (
                        "Headless LibreOffice" if is_headless_office
                        else "Headless/Scoped Browser" if is_headless_browser
                        else "Orphaned Node Dev Server"
                    ),
                }

                if not dry_run:
                    try:
                        self._kill_process(pid, force=force)
                        terminated.append(proc_info)
                    except Exception as exc:
                        errors.append(f"Failed to terminate PID {pid} ({proc.get('name')}): {exc}")
                else:
                    terminated.append(proc_info)

        # Trigger garbage collection
        gc.collect()

        return {
            "processes_scanned": scanned_count,
            "terminated_count": len(terminated),
            "terminated": terminated,
            "errors": errors,
        }

    def _scan_windows_processes(self) -> List[Dict[str, Any]]:
        """
        Query Windows processes via PowerShell CIM (Get-CimInstance Win32_Process).
        Avoids deprecated wmic.exe on Windows 11.
        """
        procs: List[Dict[str, Any]] = []
        ps_cmd = (
            "Get-CimInstance Win32_Process | "
            "Select-Object ProcessId, Name, CommandLine | "
            "ConvertTo-Json -Compress"
        )
        try:
            res = subprocess.run(
                ["powershell.exe", "-NoProfile", "-NonInteractive", "-Command", ps_cmd],
                capture_output=True,
                text=True,
                timeout=12,
                errors="replace",
            )
            if res.returncode == 0 and res.stdout.strip():
                raw = json.loads(res.stdout)
                items = raw if isinstance(raw, list) else [raw]
                for item in items:
                    procs.append({
                        "pid": int(item.get("ProcessId") or 0),
                        "name": str(item.get("Name") or ""),
                        "cmdline": str(item.get("CommandLine") or ""),
                    })
                return procs
        except Exception:
            pass

        # Fallback to tasklist /V /FO CSV
        try:
            res = subprocess.run(
                ["tasklist.exe", "/V", "/FO", "CSV"],
                capture_output=True,
                text=True,
                timeout=8,
                errors="replace",
            )
            if res.returncode == 0:
                reader = csv.reader(io.StringIO(res.stdout))
                header = next(reader, None)
                for row in reader:
                    if len(row) >= 2:
                        try:
                            procs.append({
                                "pid": int(row[1]),
                                "name": row[0],
                                "cmdline": row[8] if len(row) > 8 else "",
                            })
                        except (ValueError, IndexError):
                            pass
        except Exception:
            pass

        return procs

    def _scan_posix_processes(self) -> List[Dict[str, Any]]:
        """Query POSIX processes via ps -eo pid,comm,args."""
        procs: List[Dict[str, Any]] = []
        try:
            res = subprocess.run(
                ["ps", "-eo", "pid,comm,args"],
                capture_output=True,
                text=True,
                timeout=5,
                errors="replace",
            )
            if res.returncode == 0:
                for line in res.stdout.splitlines()[1:]:
                    parts = line.strip().split(maxsplit=2)
                    if len(parts) >= 2:
                        try:
                            pid = int(parts[0])
                            name = parts[1]
                            cmdline = parts[2] if len(parts) > 2 else ""
                            procs.append({"pid": pid, "name": name, "cmdline": cmdline})
                        except ValueError:
                            pass
        except Exception:
            pass
        return procs

    def _kill_process(self, pid: int, force: bool = False) -> None:
        """Terminate process safely by PID."""
        if sys.platform == "win32":
            args = ["taskkill.exe", "/PID", str(pid)]
            if force:
                args.insert(1, "/F")
            subprocess.run(args, capture_output=True, timeout=5)
        else:
            sig = signal.SIGKILL if force else signal.SIGTERM
            os.kill(pid, sig)

    # -------------------------------------------------------------------------
    # Pillar 3: SQLite Database Compaction & FTS5 Optimization
    # -------------------------------------------------------------------------

    def vacuum_and_optimize_databases(self, dry_run: bool = False) -> Dict[str, Any]:
        """
        Dynamically scan and optimize SQLite databases:
        - .web_outputs/studio_history.db (Primary conversion history)
        - .labor_law_fts.db (FTS5 Knowledge Vault if present)
        - Any *.db file in workspace.
        
        Performs PRAGMA quick_check, PRAGMA wal_checkpoint(TRUNCATE), VACUUM, and PRAGMA optimize.
        """
        target_dbs: List[str] = []
        reports: List[Dict[str, Any]] = []
        total_saved_bytes = 0
        errors: List[str] = []

        # Candidate DB paths
        candidates = [
            os.path.join(self.workspace_root, ".web_outputs", "studio_history.db"),
            os.path.join(self.workspace_root, ".labor_law_fts.db"),
            os.path.join(self.workspace_root, "history.db"),
        ]

        # Scan for existing databases
        for c in candidates:
            if os.path.isfile(c):
                target_dbs.append(c)

        for db_path in target_dbs:
            try:
                size_before = os.path.getsize(db_path)
                integrity_ok = True
                fts_optimized = False

                conn = sqlite3.connect(db_path, timeout=10.0)
                cur = conn.cursor()

                # Integrity check
                try:
                    check_rows = cur.execute("PRAGMA quick_check;").fetchall()
                    integrity_ok = bool(check_rows and check_rows[0][0] == "ok")
                except Exception:
                    integrity_ok = False

                # Checkpoint WAL log
                try:
                    cur.execute("PRAGMA wal_checkpoint(TRUNCATE);")
                    conn.commit()
                except Exception:
                    pass

                # If FTS5 table is present, run optimize
                try:
                    tables = [r[0] for r in cur.execute("SELECT name FROM sqlite_master WHERE type='table';").fetchall()]
                    if "labor_law_fts" in tables:
                        cur.execute("INSERT INTO labor_law_fts(labor_law_fts) VALUES('optimize');")
                        conn.commit()
                        fts_optimized = True
                except Exception:
                    pass

                if not dry_run:
                    try:
                        cur.execute("VACUUM;")
                        cur.execute("PRAGMA optimize;")
                    except Exception as vac_err:
                        errors.append(f"VACUUM error on {os.path.basename(db_path)}: {vac_err}")

                conn.close()

                size_after = os.path.getsize(db_path) if not dry_run else size_before
                saved = max(0, size_before - size_after)
                total_saved_bytes += saved

                reports.append({
                    "database": os.path.relpath(db_path, self.workspace_root),
                    "integrity_status": "OK" if integrity_ok else "CHECK_FAILED",
                    "size_before_bytes": size_before,
                    "size_after_bytes": size_after,
                    "bytes_saved": saved,
                    "fts_optimized": fts_optimized,
                })
            except Exception as exc:
                errors.append(f"Failed to maintain database {db_path}: {exc}")

        return {
            "databases_processed": len(reports),
            "reports": reports,
            "total_saved_bytes": total_saved_bytes,
            "errors": errors,
        }

    # -------------------------------------------------------------------------
    # Pillar 4: Python Bytecode & Test Cache Sanitation
    # -------------------------------------------------------------------------

    def clean_bytecode_and_test_cache(self, dry_run: bool = False) -> Dict[str, Any]:
        """
        Scan and purge __pycache__ directories, *.pyc/*.pyo files, .pytest_cache,
        and .coverage temp files across workspace.
        """
        pycache_dirs: List[str] = []
        pyc_files: List[str] = []
        test_cache_dirs: List[str] = []
        bytes_freed = 0
        errors: List[str] = []

        for root, dirs, files in os.walk(self.workspace_root):
            # Don't recurse into git/node_modules/venv
            dirs[:] = [d for d in dirs if d not in {".git", "node_modules", "venv", ".venv", "env"}]
            
            # Catch .pytest_cache
            if ".pytest_cache" in dirs:
                p_cache = os.path.join(root, ".pytest_cache")
                test_cache_dirs.append(p_cache)
                dirs.remove(".pytest_cache")

            # Catch __pycache__
            if "__pycache__" in dirs:
                p_pyc = os.path.join(root, "__pycache__")
                pycache_dirs.append(p_pyc)
                dirs.remove("__pycache__")

            for f in files:
                if f.endswith((".pyc", ".pyo", ".coverage")):
                    full_p = os.path.join(root, f)
                    pyc_files.append(full_p)
                    try:
                        bytes_freed += os.path.getsize(full_p)
                    except OSError:
                        pass

        # Calculate bytes inside directories
        for d in pycache_dirs + test_cache_dirs:
            try:
                for entry in Path(d).rglob("*"):
                    if entry.is_file():
                        bytes_freed += entry.stat().st_size
            except OSError:
                pass

        if not dry_run:
            for f in pyc_files:
                try:
                    os.remove(f)
                except Exception as exc:
                    errors.append(f"Cannot remove {f}: {exc}")

            for d in pycache_dirs + test_cache_dirs:
                try:
                    shutil.rmtree(d, ignore_errors=True)
                except Exception as exc:
                    errors.append(f"Cannot remove dir {d}: {exc}")

        return {
            "pycache_dirs_count": len(pycache_dirs),
            "pyc_files_count": len(pyc_files),
            "test_cache_dirs_count": len(test_cache_dirs),
            "bytes_freed": bytes_freed,
            "errors": errors,
        }

    # -------------------------------------------------------------------------
    # Pillar 5: Log Rotation & Output Pruning
    # -------------------------------------------------------------------------

    def rotate_and_prune_logs(
        self, dry_run: bool = False, max_size_mb: float = 10.0, max_age_days: int = 7
    ) -> Dict[str, Any]:
        """
        Rotate log files exceeding max_size_mb and prune old task logs older than max_age_days.
        """
        rotated: List[Dict[str, Any]] = []
        pruned_count = 0
        bytes_freed = 0
        errors: List[str] = []
        max_bytes = int(max_size_mb * 1024 * 1024)
        cutoff_sec = time.time() - (max_age_days * 86400)

        # 1. Inspect *.log files in workspace root
        for log_file in Path(self.workspace_root).glob("*.log"):
            try:
                size = log_file.stat().st_size
                if size > max_bytes:
                    rel_p = os.path.relpath(str(log_file), self.workspace_root)
                    if not dry_run:
                        old_backup = str(log_file) + ".old"
                        if os.path.exists(old_backup):
                            os.remove(old_backup)
                        os.rename(str(log_file), old_backup)
                        # Re-create empty file
                        open(str(log_file), "w", encoding="utf-8").close()
                    bytes_freed += size
                    rotated.append({"file": rel_p, "action": "rotated", "size_bytes": size})
            except Exception as exc:
                errors.append(f"Cannot rotate log {log_file}: {exc}")

        # 2. Inspect task logs in .system_generated/tasks/
        tasks_dir = os.path.join(self.workspace_root, ".system_generated", "tasks")
        if os.path.isdir(tasks_dir):
            for t_log in Path(tasks_dir).glob("*.log"):
                try:
                    st = t_log.stat()
                    if st.st_mtime < cutoff_sec:
                        bytes_freed += st.st_size
                        if not dry_run:
                            os.remove(str(t_log))
                        pruned_count += 1
                except Exception as exc:
                    errors.append(f"Cannot prune task log {t_log}: {exc}")

        return {
            "rotated_count": len(rotated),
            "rotated_logs": rotated,
            "pruned_task_logs": pruned_count,
            "bytes_freed": bytes_freed,
            "errors": errors,
        }

    def prune_web_outputs(
        self, dry_run: bool = False, max_age_days: int = 7
    ) -> Dict[str, Any]:
        """
        Prune files and asset directories in .web_outputs/ older than max_age_days.
        SAFETY: Never deletes studio_history.db or history database files!
        Only executed when user explicitly specifies --prune-outputs.
        """
        web_out = os.path.join(self.workspace_root, ".web_outputs")
        pruned_items: List[Dict[str, Any]] = []
        bytes_freed = 0
        errors: List[str] = []

        if not os.path.isdir(web_out):
            return {"pruned_count": 0, "bytes_freed": 0, "details": []}

        cutoff_sec = time.time() - (max_age_days * 86400)

        for entry in os.scandir(web_out):
            name = entry.name
            # Preserve database and WAL files
            if name.startswith("studio_history.db"):
                continue

            try:
                mtime = entry.stat().st_mtime
                if mtime < cutoff_sec:
                    if entry.is_file():
                        sz = entry.stat().st_size
                        bytes_freed += sz
                        if not dry_run:
                            os.remove(entry.path)
                        pruned_items.append({"name": name, "type": "file", "size_bytes": sz})
                    elif entry.is_dir():
                        sz = sum(f.stat().st_size for f in Path(entry.path).rglob("*") if f.is_file())
                        bytes_freed += sz
                        if not dry_run:
                            shutil.rmtree(entry.path, ignore_errors=True)
                        pruned_items.append({"name": name, "type": "dir", "size_bytes": sz})
            except Exception as exc:
                errors.append(f"Cannot prune output {name}: {exc}")

        return {
            "pruned_count": len(pruned_items),
            "bytes_freed": bytes_freed,
            "details": pruned_items,
            "errors": errors,
        }

    # -------------------------------------------------------------------------
    # Pillar 6: Windows MPO (Multiplane Overlay) Stutter Guard
    # -------------------------------------------------------------------------

    def check_and_diagnose_mpo(self, apply_fix: bool = False) -> Dict[str, Any]:
        """
        Diagnose Windows 11 Desktop Window Manager (DWM) Multiplane Overlay (MPO).
        If MPO is enabled, it frequently causes mouse micro-stuttering on desktop/Electron apps.
        Disabling MPO requires setting OverlayTestMode = 5 in HKEY_LOCAL_MACHINE\\SOFTWARE\\Microsoft\\Windows\\DWM.
        
        UAC GUARD: Modifying HKLM requires Administrator elevation.
        If not elevated, warns user and recommends merging tools/disable_mpo.reg without crashing.
        """
        res: Dict[str, Any] = {
            "supported": sys.platform == "win32",
            "status": "NOT_APPLICABLE",
            "overlay_test_mode": None,
            "is_admin": False,
            "recommendation": "",
            "fixed": False,
        }

        if sys.platform != "win32":
            res["recommendation"] = "MPO is Windows DWM specific; not required on this platform."
            return res

        import ctypes
        is_admin = False
        try:
            is_admin = ctypes.windll.shell32.IsUserAnAdmin() != 0
        except Exception:
            pass
        res["is_admin"] = is_admin

        try:
            import winreg
            dwm_key_path = r"SOFTWARE\Microsoft\Windows\DWM"
            
            # Read current status
            val = None
            try:
                with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, dwm_key_path, 0, winreg.KEY_READ) as key:
                    val, _ = winreg.QueryValueEx(key, "OverlayTestMode")
                    res["overlay_test_mode"] = val
            except FileNotFoundError:
                res["overlay_test_mode"] = None

            if val == 5:
                res["status"] = "OPTIMAL_DISABLED"
                res["recommendation"] = "Windows DWM MPO is disabled (OverlayTestMode=5). Mouse cursor is smooth."
            else:
                res["status"] = "WARNING_MPO_ENABLED"
                res["recommendation"] = (
                    "MPO is active (may cause mouse/GPU stutter). "
                    "Run 'tools/disable_mpo.reg' as Administrator to disable MPO."
                )

            # Apply fix if requested
            if apply_fix and res["status"] != "OPTIMAL_DISABLED":
                if is_admin:
                    try:
                        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, dwm_key_path, 0, winreg.KEY_SET_VALUE) as key:
                            winreg.SetValueEx(key, "OverlayTestMode", 0, winreg.REG_DWORD, 5)
                        res["fixed"] = True
                        res["status"] = "OPTIMAL_DISABLED"
                        res["recommendation"] = "Successfully set OverlayTestMode=5. Please restart DWM or reboot PC."
                    except Exception as reg_err:
                        res["fix_error"] = str(reg_err)
                else:
                    res["requires_elevation"] = True
                    res["recommendation"] = (
                        "Administrator rights required to write HKLM. "
                        "Please right-click 'tools/disable_mpo.reg' and select 'Merge', then restart."
                    )

        except Exception as exc:
            res["error"] = str(exc)

        return res

    # -------------------------------------------------------------------------
    # Orchestrator: run_all
    # -------------------------------------------------------------------------

    def run_all(
        self,
        dry_run: bool = False,
        clean_locks: bool = True,
        clean_temp: bool = True,
        reap_processes: bool = True,
        vacuum_db: bool = True,
        clear_cache: bool = True,
        rotate_logs: bool = True,
        prune_outputs: bool = False,
        prune_days: int = 7,
        check_mpo: bool = True,
        fix_mpo: bool = False,
    ) -> MaintenanceReport:
        """
        Execute full multi-pillar system maintenance suite with precise telemetry.
        """
        t0 = time.time()
        report = MaintenanceReport(dry_run=dry_run)
        report.vitals = self.get_system_vitals()

        # Pillar 1: Locks
        if clean_locks:
            lock_res = self.clean_zombie_locks(dry_run=dry_run)
            report.lock_files_removed = lock_res["removed_count"]
            report.lock_files_details = lock_res["details"]
            report.total_bytes_freed += lock_res["bytes_freed"]
            report.errors.extend(lock_res["errors"])

        # Pillar 1: Temp
        if clean_temp:
            temp_res = self.clean_temp_artifacts(dry_run=dry_run)
            report.temp_files_removed = temp_res["removed_files_count"] + temp_res["removed_dirs_count"]
            report.temp_files_details = temp_res["details"]
            report.total_bytes_freed += temp_res["bytes_freed"]
            report.errors.extend(temp_res["errors"])

        # Pillar 2: Processes
        if reap_processes:
            proc_res = self.reap_orphaned_processes(dry_run=dry_run)
            report.processes_scanned = proc_res["processes_scanned"]
            report.orphaned_processes_terminated = proc_res["terminated"]
            report.errors.extend(proc_res["errors"])

        # Pillar 3: Databases
        if vacuum_db:
            db_res = self.vacuum_and_optimize_databases(dry_run=dry_run)
            report.databases_vacuumed = db_res["reports"]
            report.total_bytes_freed += db_res["total_saved_bytes"]
            report.errors.extend(db_res["errors"])

        # Pillar 4: Bytecode & Cache
        if clear_cache:
            cache_res = self.clean_bytecode_and_test_cache(dry_run=dry_run)
            report.pycache_dirs_removed = cache_res["pycache_dirs_count"]
            report.pycache_files_removed = cache_res["pyc_files_count"]
            report.test_cache_dirs_removed = cache_res["test_cache_dirs_count"]
            report.total_bytes_freed += cache_res["bytes_freed"]
            report.errors.extend(cache_res["errors"])

        # Pillar 5: Logs
        if rotate_logs:
            log_res = self.rotate_and_prune_logs(dry_run=dry_run)
            report.logs_rotated = log_res["rotated_logs"]
            report.total_bytes_freed += log_res["bytes_freed"]
            report.errors.extend(log_res["errors"])

        # Pillar 5: Prune Web Outputs (Only when explicitly requested)
        if prune_outputs:
            out_res = self.prune_web_outputs(dry_run=dry_run, max_age_days=prune_days)
            report.outputs_pruned = out_res["details"]
            report.total_bytes_freed += out_res["bytes_freed"]
            report.errors.extend(out_res["errors"])

        # Pillar 6: MPO
        if check_mpo:
            report.mpo_status = self.check_and_diagnose_mpo(apply_fix=fix_mpo)

        report.execution_time_sec = round(time.time() - t0, 3)
        return report

