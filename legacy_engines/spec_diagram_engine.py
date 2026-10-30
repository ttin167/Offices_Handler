# -*- coding: utf-8 -*-
"""
legacy_engines/spec_diagram_engine.py — Thin Backward-Compatibility Shim.
Redirects to core.diagrams.legacy.canvas (Single Source of Truth).
"""
from core.diagrams.legacy.canvas import *  # noqa: F401,F403
from core.diagrams.legacy.canvas import (
    PrecisionDiagram,
    Node,
    Edge,
    NodeType,
    PortType,
    DiagramStyle,
    PALETTE,
    _find_chromium_executable,
)

__all__ = [
    "PrecisionDiagram",
    "Node",
    "Edge",
    "NodeType",
    "PortType",
    "DiagramStyle",
    "PALETTE",
    "_find_chromium_executable",
]
