"""
ui — Desktop GUI Interface for Antigravity Universal Document Studio.
"""

from __future__ import annotations

from .desktop_app import UniversalDocStudioApp, launch_gui
from .webview_launcher import (
    build_app_mode_browser_command,
    find_app_mode_browser,
    launch_native_web_studio,
)

__all__ = [
    "UniversalDocStudioApp",
    "launch_gui",
    "launch_native_web_studio",
    "build_app_mode_browser_command",
    "find_app_mode_browser",
]
