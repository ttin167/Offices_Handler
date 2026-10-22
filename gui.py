"""
gui.py — Legacy Desktop GUI Facade.

Thin Facade pointing to ui.desktop_app for strict Layered Modular Architecture.
"""
from __future__ import annotations

import os
import sys

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)

from ui.desktop_app import *
import ui.desktop_app as _mod

for _k in dir(_mod):
    if not _k.startswith("__"):
        globals()[_k] = getattr(_mod, _k)


if __name__ == "__main__":
    from ui.desktop_app import launch_gui
    launch_gui()
