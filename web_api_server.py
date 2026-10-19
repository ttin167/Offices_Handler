"""
web_api_server.py — High-Performance Zero-Dependency Python Web API Bridge Facade.

Thin Facade pointing to server.app for strict Layered Modular Architecture.
Maintains 100% backward compatibility for all legacy callers and tests.
"""
from __future__ import annotations

import os
import sys
from server.app import (
    OfficeApiHandler,
    run_server,
    is_server_already_running,
    export_pptx_slides,
    cleanup_old_temp_files,
    HOST,
    PORT,
    BASE_DIR,
    DIST_DIR,
    UPLOAD_DIR,
    OUTPUT_DIR,
)

__all__ = [
    "OfficeApiHandler",
    "run_server",
    "is_server_already_running",
    "export_pptx_slides",
    "cleanup_old_temp_files",
    "HOST",
    "PORT",
    "BASE_DIR",
    "DIST_DIR",
    "UPLOAD_DIR",
    "OUTPUT_DIR",
]

if __name__ == "__main__":
    port_arg = int(sys.argv[1]) if len(sys.argv) > 1 and sys.argv[1].isdigit() else int(os.environ.get("PORT", PORT))
    host_arg = sys.argv[2] if len(sys.argv) > 2 else os.environ.get("HOST", HOST)
    run_server(port_arg, host_arg)
