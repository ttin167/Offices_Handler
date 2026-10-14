"""
core.diagrams — Technical Architecture, ERD, and Diagram Generation Suite.
"""

from __future__ import annotations

from .mxgraph import render_diagram
from .editor import DiagramEditorApp

__all__ = [
    "render_diagram",
    "DiagramEditorApp",
]
