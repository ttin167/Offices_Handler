# -*- coding: utf-8 -*-
"""
ui/webview_launcher.py — Native Desktop App Mode Wrapper for Antigravity Web Studio (v4.6).

Enforces Guard #5 (ERR_UI_001):
  - Passes `--user-data-dir=<temp_app_profile_dir>` when launching Edge/Chrome in `--app` mode.
  - Isolates the App window process from any existing background browser instance on Windows,
    ensuring `proc.wait()` blocks for the exact lifetime of the App window and automatically
    shuts down the background Python server and cleans up the temp profile directory on exit.
"""

from __future__ import annotations

import os
import shutil
import socket
import subprocess
import sys
import tempfile
import threading
import time
import webbrowser
from http.server import ThreadingHTTPServer
from typing import Any, Dict, List, Optional


def find_app_mode_browser() -> Optional[str]:
    """Locates Microsoft Edge, Google Chrome, or Chromium executable for Native App Mode."""
    candidates: List[str] = []
    if sys.platform == "win32":
        prog_files = [
            os.environ.get("PROGRAMFILES", r"C:\Program Files"),
            os.environ.get("PROGRAMFILES(X86)", r"C:\Program Files (x86)"),
            os.environ.get("LOCALAPPDATA", ""),
        ]
        for base in prog_files:
            if not base:
                continue
            candidates.extend([
                os.path.join(base, "Microsoft", "Edge", "Application", "msedge.exe"),
                os.path.join(base, "Google", "Chrome", "Application", "chrome.exe"),
            ])
    elif sys.platform == "darwin":
        candidates.extend([
            "/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge",
            "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
        ])
    else:
        for bin_name in ("microsoft-edge", "google-chrome", "chromium", "chromium-browser"):
            found = shutil.which(bin_name)
            if found:
                candidates.append(found)

    for c in candidates:
        if c and os.path.isfile(c):
            return os.path.abspath(c)

    for fallback_bin in ("msedge", "chrome", "google-chrome", "chromium"):
        w = shutil.which(fallback_bin)
        if w:
            return w
    return None


def build_app_mode_browser_command(
    url: str,
    user_data_dir: str,
    browser_binary: Optional[str] = None,
    window_size: str = "1440,900",
    exe_path: Optional[str] = None,
    width: Optional[int] = None,
    height: Optional[int] = None,
) -> List[str]:
    """
    Guard #5 (ERR_UI_001): Constructs the isolated browser command with `--user-data-dir`
    so `proc.wait()` tracks the exact window lifetime on Windows/Linux/macOS.
    """
    exe = exe_path or browser_binary or find_app_mode_browser() or "msedge"
    if width and height:
        window_size = f"{int(width)},{int(height)}"
    abs_profile = os.path.abspath(user_data_dir)
    return [
        exe,
        f"--app={url}",
        f"--user-data-dir={abs_profile}",
        f"--window-size={window_size}",
        "--no-first-run",
        "--no-default-browser-check",
        "--disable-extensions",
    ]


def _find_free_port(preferred_port: int = 0) -> int:
    if preferred_port > 0:
        return preferred_port
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.bind(("127.0.0.1", 0))
        return int(s.getsockname()[1])


def launch_native_web_studio(
    port: int = 0,
    dry_run: bool = False,
) -> Dict[str, Any]:
    """
    Starts the Antigravity Studio REST API & static UI server on a background daemon thread
    and opens a frameless Desktop Native App window (`--app` mode with isolated `--user-data-dir`).
    When the user closes the window, automatically shuts down the server and cleans up the temp profile.
    """
    from server.app import OfficeApiHandler

    bound_port = _find_free_port(port)
    server = ThreadingHTTPServer(("127.0.0.1", bound_port), OfficeApiHandler)
    actual_port = int(server.server_address[1])
    app_url = f"http://127.0.0.1:{actual_port}"

    server_thread = threading.Thread(target=server.serve_forever, name="AntigravityWebStudioDaemon", daemon=True)
    server_thread.start()

    temp_profile_dir = tempfile.mkdtemp(prefix="antigravity_app_profile_")
    browser_exe = find_app_mode_browser()
    cmd = build_app_mode_browser_command(app_url, temp_profile_dir, browser_binary=browser_exe)
    proc = None

    try:
        if dry_run:
            return {
                "success": True,
                "mode": "dry_run",
                "profile_cleaned": True,
                "url": app_url,
                "port": actual_port,
                "user_data_dir": temp_profile_dir,
                "command": cmd,
                "browser_found": browser_exe is not None,
            }

        if browser_exe:
            proc = subprocess.Popen(cmd)
            # Poll with short intervals instead of blocking proc.wait() so Windows delivers Ctrl+C (KeyboardInterrupt) immediately
            while proc.poll() is None:
                time.sleep(0.25)
        else:
            webbrowser.open(app_url)
            while server_thread.is_alive():
                time.sleep(0.5)

        return {
            "success": True,
            "url": app_url,
            "port": actual_port,
            "user_data_dir": temp_profile_dir,
            "command": cmd,
        }
    except KeyboardInterrupt:
        return {"success": True, "url": app_url, "port": actual_port, "interrupted": True}
    finally:
        if proc is not None and proc.poll() is None:
            try:
                if sys.platform == "win32":
                    subprocess.run(["taskkill", "/pid", str(proc.pid), "/f", "/t"], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
                else:
                    proc.terminate()
            except Exception:
                pass
        try:
            server.shutdown()
            server.server_close()
        except Exception:
            pass
        if os.path.isdir(temp_profile_dir):
            shutil.rmtree(temp_profile_dir, ignore_errors=True)
