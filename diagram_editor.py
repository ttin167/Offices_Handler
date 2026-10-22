"""
diagram_editor.py — Interactive Canvas Editor Facade.

Thin Facade pointing to core.diagrams.editor for strict Layered Modular Architecture.
"""
from __future__ import annotations

import os
import sys

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)

from core.diagrams.editor import *
import core.diagrams.editor as _mod

for _k in dir(_mod):
    if not _k.startswith("__"):
        globals()[_k] = getattr(_mod, _k)


if __name__ == "__main__":
    from core.diagrams.editor import main as _editor_main
    _editor_main()
