"""
main.py — Main Application Entry Point for Universal Office Studio.

Auto-routes:
- No arguments: Launches Desktop Tkinter GUI (ui.desktop_app)
- With CLI arguments: Dispatches to Unified AI Tools CLI (cli.ai_tools)
"""
from __future__ import annotations

import os
import sys

if sys.platform == "win32":
    try:
        if hasattr(sys.stdout, "reconfigure"):
            getattr(sys.stdout, "reconfigure")(encoding="utf-8", errors="replace")
        if hasattr(sys.stderr, "reconfigure"):
            getattr(sys.stderr, "reconfigure")(encoding="utf-8", errors="replace")
    except Exception:
        pass

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)


def main() -> int:
    if len(sys.argv) > 1:
        if sys.argv[1] in ("--web-app", "--app"):
            from ui.webview_launcher import launch_native_web_studio
            launch_native_web_studio()
            return 0
        from cli.ai_tools import main as run_cli
        return run_cli()
    else:
        from ui.desktop_app import launch_gui
        launch_gui()
        return 0


if __name__ == "__main__":
    sys.exit(main())
