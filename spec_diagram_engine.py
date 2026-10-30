"""
spec_diagram_engine.py — SVG Canvas Precision Diagram CLI & Engine Entrypoint.

Thin Facade pointing to core.diagrams.legacy.canvas for Layered Modular Architecture.
"""

from __future__ import annotations

import os
import sys

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)

from core.diagrams.legacy.canvas import *  # type: ignore
from core.diagrams.legacy.canvas import (
    PrecisionDiagram,
    Node,
    Edge,
    NodeType,
    PortType,
    ColumnDef,
    ClusterDef,
    _find_chromium_executable,
)

if __name__ == "__main__":
    from core.diagrams.legacy.canvas import main as _canvas_main
    _canvas_main()
