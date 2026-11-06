# -*- coding: utf-8 -*-
"""
server.routes.system_routes — Health, Environment Doctor & System Maintenance Endpoints.
"""

from __future__ import annotations

import os
import sys
import time
from typing import Any

from .base import router, send_json, send_error, read_json_body


@router.get("/api/health")
def handle_health(handler: Any) -> None:
    """Lightweight health check endpoint for container probes."""
    send_json(handler, {
        "status": "healthy",
        "service": "antigravity-doc-handler",
        "version": "5.0.0",
        "timestamp": int(time.time()),
    })


@router.get("/api/system/vitals")
def handle_system_vitals(handler: Any) -> None:
    """Returns system memory, disk, and runtime vitals."""
    try:
        from core.system_maintenance import get_system_vitals
        vitals = get_system_vitals()
        send_json(handler, {"success": True, "vitals": vitals})
    except Exception as e:
        send_error(handler, f"Lỗi lấy thông số hệ thống: {e}", status=500)


@router.post("/api/system/maintenance")
def handle_system_maintenance(handler: Any) -> None:
    """Triggers 7-pillar system maintenance & cleanup."""
    body = read_json_body(handler)
    dry_run = body.get("dry_run", False)
    try:
        from core.system_maintenance import run_system_maintenance
        report = run_system_maintenance(dry_run=dry_run)
        from dataclasses import asdict
        send_json(handler, {"success": True, "report": asdict(report)})
    except Exception as e:
        send_error(handler, f"Lỗi chạy bảo trì hệ thống: {e}", status=500)

