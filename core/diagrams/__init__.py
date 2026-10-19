"""
core.diagrams — Technical Architecture, ERD, and Diagram Generation Suite.
"""

from __future__ import annotations

from .mxgraph import render_diagram
from .editor import DiagramEditorApp
from .bridge import (
    calculate_aspect_dimensions,
    render_diagram_code_to_png,
    render_diagram_file_to_png,
    DIAGRAM_LANGUAGES,
    DIAGRAM_EXTENSIONS,
    IMAGE_EXTENSIONS,
)

__all__ = [
    "render_diagram",
    "DiagramEditorApp",
    "calculate_aspect_dimensions",
    "render_diagram_code_to_png",
    "render_diagram_file_to_png",
    "DIAGRAM_LANGUAGES",
    "DIAGRAM_EXTENSIONS",
    "IMAGE_EXTENSIONS",
]
