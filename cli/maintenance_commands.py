"""
cli.maintenance_commands — System Maintenance CLI Subcommand Dispatcher.

Handles `python ai_tools_cli.py system-maintenance [options]` to trigger
periodic multi-pillar workspace maintenance, resource hygiene, and diagnostics.
"""

from __future__ import annotations

import argparse
import json
import os
import sys
import time
from typing import Any, Dict

from core.system_maintenance import SystemMaintenanceEngine, MaintenanceReport


def _format_bytes(num_bytes: int) -> str:
    """Format bytes into human-readable string (B, KB, MB, GB)."""
    if num_bytes < 1024:
        return f"{num_bytes} B"
    elif num_bytes < 1024 ** 2:
        return f"{num_bytes / 1024:.1f} KB"
    elif num_bytes < 1024 ** 3:
        return f"{num_bytes / (1024 ** 2):.1f} MB"
    else:
        return f"{num_bytes / (1024 ** 3):.2f} GB"


def cmd_system_maintenance(args: argparse.Namespace) -> int:
    """Execute system maintenance command line."""
    engine = SystemMaintenanceEngine()
    
    # Determine selective modules
    is_selective = any([
        args.clean_locks,
        args.clean_temp,
        args.reap_processes,
        args.vacuum_db,
        args.clear_cache,
        args.rotate_logs,
        args.prune_outputs,
        args.mpo_check,
        args.fix_mpo,
    ])

    if is_selective:
        clean_locks = args.clean_locks
        clean_temp = args.clean_temp
        reap_processes = args.reap_processes
        vacuum_db = args.vacuum_db
        clear_cache = args.clear_cache
        rotate_logs = args.rotate_logs
        prune_outputs = args.prune_outputs
        check_mpo = args.mpo_check or args.fix_mpo
    else:
        # Default full maintenance (outputs pruning is explicitly opt-in)
        clean_locks = True
        clean_temp = True
        reap_processes = True
        vacuum_db = True
        clear_cache = True
        rotate_logs = True
        prune_outputs = args.prune_outputs
        check_mpo = True

    fix_mpo = bool(args.fix_mpo)
    dry_run = bool(args.dry_run)

    report = engine.run_all(
        dry_run=dry_run,
        clean_locks=clean_locks,
        clean_temp=clean_temp,
        reap_processes=reap_processes,
        vacuum_db=vacuum_db,
        clear_cache=clear_cache,
        rotate_logs=rotate_logs,
        prune_outputs=prune_outputs,
        prune_days=args.days,
        check_mpo=check_mpo,
        fix_mpo=fix_mpo,
    )

    if args.json:
        print(json.dumps(report.to_dict(), indent=2, ensure_ascii=False))
        return 0 if not report.errors else 1

    # Human-Readable ANSI Terminal Output
    print("\n" + "=" * 76)
    mode_str = "[DRY-RUN PREVIEW]" if dry_run else "[LIVE MAINTENANCE]"
    print(f" {mode_str} ANTIGRAVITY SYSTEM MAINTENANCE SUITE (v4.6.0)")
    print("=" * 76)
    
    v = report.vitals
    print(f"OS Platform     : {v.get('os_platform')} ({v.get('os_release', 'N/A')})")
    print(f"Python Runtime  : {v.get('python_version')} | Execution: {report.execution_time_sec:.3f}s")
    if "disk_free_gb" in v:
        print(f"Disk Capacity   : {v.get('disk_free_gb')} GB free / {v.get('disk_total_gb')} GB total ({v.get('disk_free_pct')}%)")
    if "ram_available_gb" in v:
        print(f"Memory (RAM)    : {v.get('ram_available_gb')} GB available / {v.get('ram_total_gb')} GB total (Load: {v.get('ram_load_pct')}%)")
    print("-" * 76)

    # Pillar 1
    if clean_locks or clean_temp:
        print(f"[*] Pillar 1: Lock Files & Temp Leaks (ERR_CONV_004):")
        print(f"    - Zombie Locks Removed : {report.lock_files_removed}")
        print(f"    - Temp Artifacts Purged: {report.temp_files_removed}")

    # Pillar 2
    if reap_processes:
        print(f"[*] Pillar 2: Process Reaper & Memory Hygiene:")
        print(f"    - Processes Scanned    : {report.processes_scanned}")
        print(f"    - Orphaned Terminated  : {len(report.orphaned_processes_terminated)}")
        for p in report.orphaned_processes_terminated:
            print(f"      + [{p.get('pid')}] {p.get('name')} -> {p.get('reason')}")

    # Pillar 3
    if vacuum_db:
        print(f"[*] Pillar 3: SQLite Database Compaction & FTS5:")
        for db in report.databases_vacuumed:
            diff = _format_bytes(db.get('bytes_saved', 0))
            fts_note = " [FTS5 Optimized]" if db.get('fts_optimized') else ""
            print(f"    - {db.get('database')}: {db.get('integrity_status')} | Saved: {diff}{fts_note}")

    # Pillar 4
    if clear_cache:
        print(f"[*] Pillar 4: Bytecode & Test Cache Sanitation:")
        print(f"    - __pycache__ Dirs     : {report.pycache_dirs_removed}")
        print(f"    - .pyc Bytecode Files  : {report.pycache_files_removed}")
        print(f"    - .pytest_cache Dirs   : {report.test_cache_dirs_removed}")

    # Pillar 5
    if rotate_logs or prune_outputs:
        print(f"[*] Pillar 5: Log Rotation & Output Storage:")
        print(f"    - Logs Rotated         : {len(report.logs_rotated)}")
        if prune_outputs:
            print(f"    - Outputs Pruned (>={args.days}d): {len(report.outputs_pruned)}")

    # Pillar 6
    if check_mpo:
        mpo = report.mpo_status
        print(f"[*] Pillar 6: Windows MPO Desktop Stutter Guard:")
        print(f"    - Status               : {mpo.get('status')}")
        print(f"    - Recommendation       : {mpo.get('recommendation')}")
        if mpo.get('fixed'):
            print(f"    - Remediation          : [APPLIED] OverlayTestMode set to 5.")

    print("-" * 76)
    print(f"[SUMMARY] Total Disk Space Reclaimed: {_format_bytes(report.total_bytes_freed)}")
    if report.warnings:
        print(f"[WARNINGS] {len(report.warnings)} warnings recorded.")
        for w in report.warnings:
            print(f"  ! {w}")
    if report.errors:
        print(f"[ERRORS] {len(report.errors)} errors encountered:")
        for err in report.errors:
            print(f"  x {err}")
    print("=" * 76 + "\n")

    return 0 if not report.errors else 1


def register_maintenance_subcommands(subparsers: argparse._SubParsersAction) -> None:
    """Register 'system-maintenance' CLI subcommand."""
    p_maint = subparsers.add_parser(
        "system-maintenance",
        help="Periodic system maintenance, lock file purge, orphaned process reaper, and SQLite vacuum",
        description=(
            "Automated Multi-Pillar Maintenance Suite for Antigravity Universal Office Studio.\n"
            "Purges zombie lock files (~$*), terminates headless orphaned engines, vacuums SQLite\n"
            "databases, clears bytecode cache, manages logs, and monitors Windows MPO desktop stuttering."
        ),
    )
    p_maint.add_argument(
        "--dry-run",
        action="store_true",
        help="Simulate maintenance operations without modifying files or killing processes",
    )
    p_maint.add_argument(
        "--json",
        action="store_true",
        help="Output machine-readable JSON telemetry for AI agents and automation",
    )
    p_maint.add_argument(
        "--clean-locks",
        action="store_true",
        help="Pillar 1: Purge zombie lock files (~$*.docx, .~lock.*#)",
    )
    p_maint.add_argument(
        "--clean-temp",
        action="store_true",
        help="Pillar 1: Purge temporary pipeline files (step_*_temp.*, scoped_dir*)",
    )
    p_maint.add_argument(
        "--reap-processes",
        action="store_true",
        help="Pillar 2: Terminate orphaned headless processes (soffice, headless chromium)",
    )
    p_maint.add_argument(
        "--vacuum-db",
        action="store_true",
        help="Pillar 3: Vacuum and checkpoint SQLite databases (studio_history.db, FTS5)",
    )
    p_maint.add_argument(
        "--clear-cache",
        action="store_true",
        help="Pillar 4: Remove __pycache__, *.pyc, and .pytest_cache directories",
    )
    p_maint.add_argument(
        "--rotate-logs",
        action="store_true",
        help="Pillar 5: Rotate logs exceeding 10MB and prune old task logs",
    )
    p_maint.add_argument(
        "--prune-outputs",
        action="store_true",
        help="Pillar 5: Prune old generated outputs in .web_outputs/ (opt-in)",
    )
    p_maint.add_argument(
        "--days",
        type=int,
        default=7,
        help="Retention period in days for output pruning (default: 7)",
    )
    p_maint.add_argument(
        "--mpo-check",
        action="store_true",
        help="Pillar 6: Check Windows DWM Multiplane Overlay (MPO) mouse stuttering status",
    )
    p_maint.add_argument(
        "--fix-mpo",
        action="store_true",
        help="Pillar 6: Set OverlayTestMode=5 in registry (requires Administrator rights)",
    )
    p_maint.add_argument(
        "--force",
        action="store_true",
        help="Force terminate hung processes immediately",
    )
    p_maint.set_defaults(func=cmd_system_maintenance)

