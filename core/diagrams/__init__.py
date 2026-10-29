"""
core.diagrams — Technical Architecture, ERD, and Diagram Generation Suite.
"""

from __future__ import annotations

from .mxgraph import render_diagram
from .editor import DiagramEditorApp
from .code_to_diagram import (
    export_mxgraph_to_svg,
    generate_diagram_from_code,
    scan_python_ast_to_spec,
    scan_sql_or_prisma_to_erd_spec,
)
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
    "export_mxgraph_to_svg",
    "generate_diagram_from_code",
    "scan_python_ast_to_spec",
    "scan_sql_or_prisma_to_erd_spec",
    "calculate_aspect_dimensions",
    "render_diagram_code_to_png",
    "render_diagram_file_to_png",
    "DIAGRAM_LANGUAGES",
    "DIAGRAM_EXTENSIONS",
    "IMAGE_EXTENSIONS",
]
