# -*- coding: utf-8 -*-
"""
server.routes.system_routes — Health, Environment Doctor, Session Purge & System Maintenance Endpoints.
"""

from __future__ import annotations

import os
import shutil
import time
from dataclasses import asdict
from typing import Any, List

from .base import router, send_json, send_error, read_json_body


@router.get("/api/health")
def handle_health(handler: Any) -> None:
    """Lightweight local health check endpoint."""
    send_json(handler, {
        "status": "healthy",
        "service": "antigravity-doc-handler",
        "version": "5.0.0",
        "local_only": os.environ.get("ANTIGRAVITY_LOCAL_ONLY", "1").strip().lower() not in ("0", "false", "no"),
        "host": os.environ.get("HOST", "127.0.0.1"),
        "timestamp": int(time.time()),
    })


@router.get("/api/system/vitals")
def handle_system_vitals(handler: Any) -> None:
    """Returns system memory, disk, and runtime vitals."""
    try:
        from core.system_maintenance import SystemMaintenanceEngine
        vitals = SystemMaintenanceEngine().get_system_vitals()
        send_json(handler, {"success": True, "vitals": vitals})
    except Exception as e:
        send_error(handler, f"Lỗi lấy thông số hệ thống: {e}", status=500)


@router.post("/api/system/maintenance")
def handle_system_maintenance(handler: Any) -> None:
    """Triggers 7-pillar system maintenance & cleanup."""
    body = read_json_body(handler)
    try:
        from core.system_maintenance import SystemMaintenanceEngine
        engine = SystemMaintenanceEngine()
        report = engine.run_all(
            dry_run=bool(body.get("dry_run", False)),
            clean_locks=bool(body.get("clean_locks", True)),
            clean_temp=bool(body.get("clean_temp", True)),
            reap_processes=bool(body.get("reap_processes", True)),
            vacuum_db=bool(body.get("vacuum_db", True)),
            clear_cache=bool(body.get("clear_cache", True)),
            rotate_logs=bool(body.get("rotate_logs", True)),
            prune_outputs=bool(body.get("prune_outputs", False)),
            prune_days=int(body.get("prune_days", 7)),
            check_mpo=bool(body.get("check_mpo", True)),
            fix_mpo=bool(body.get("fix_mpo", False)),
        )
        send_json(handler, {"success": True, "report": asdict(report)})
    except Exception as e:
        send_error(handler, f"Lỗi chạy bảo trì hệ thống: {e}", status=500)


@router.post("/api/session/purge")
def handle_session_purge(handler: Any) -> None:
    """
    Immediately purges temporary session files (.web_uploads and .web_outputs).
    Optionally clears job history SQLite records when purge_history=True.
    """
    body = read_json_body(handler)
    purge_history = bool(body.get("purge_history", False))

    try:
        base_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
        target_dirs = [
            os.path.join(base_dir, ".web_uploads"),
            os.path.join(base_dir, ".web_outputs"),
        ]
        removed_files = 0
        bytes_freed = 0
        details: List[str] = []

        for dpath in target_dirs:
            if not os.path.isdir(dpath):
                os.makedirs(dpath, exist_ok=True)
                continue
            for entry in os.listdir(dpath):
                full_p = os.path.join(dpath, entry)
                try:
                    if os.path.isfile(full_p) or os.path.islink(full_p):
                        bytes_freed += os.path.getsize(full_p) if os.path.isfile(full_p) else 0
                        os.remove(full_p)
                        removed_files += 1
                        details.append(entry)
                    elif os.path.isdir(full_p):
                        for root, _, files in os.walk(full_p):
                            for fn in files:
                                fp = os.path.join(root, fn)
                                try:
                                    bytes_freed += os.path.getsize(fp)
                                    removed_files += 1
                                except OSError:
                                    pass
                        shutil.rmtree(full_p, ignore_errors=True)
                        details.append(f"{entry}/")
                except OSError:
                    pass

        history_cleared = False
        if purge_history:
            try:
                from server.db import clear_history
                history_cleared = bool(clear_history())
            except Exception:
                history_cleared = False

        send_json(handler, {
            "success": True,
            "removed_files": removed_files,
            "bytes_freed": bytes_freed,
            "mb_freed": round(bytes_freed / (1024 * 1024), 3),
            "history_cleared": history_cleared,
            "details": details[:50],
        })
    except Exception as e:
        send_error(handler, f"Lỗi dọn sạch dấu vết phiên: {e}", status=500)
