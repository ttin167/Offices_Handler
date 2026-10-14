"""
core.diagrams.legacy — Legacy rendering engines (PlantUML, Mermaid, Precision SVG Canvas).
Maintained for backwards compatibility. Prefer mxgraph for new diagrams.
"""

from __future__ import annotations

from .plantuml import render_plantuml_to_png
from .mermaid import render_mermaid_to_png
from .canvas import PrecisionDiagram

__all__ = [
    "render_plantuml_to_png",
    "render_mermaid_to_png",
    "PrecisionDiagram",
]
