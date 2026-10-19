"""
diagram_bridge.py — Root Facade for Unified Multi-Engine Diagram Bridge.
Re-exports symbols from core.diagrams.bridge conforming to Layered Modular Architecture.
"""

from __future__ import annotations

try:
    from core.diagrams.bridge import (
        calculate_aspect_dimensions,
        render_diagram_code_to_png,
        render_diagram_file_to_png,
        DIAGRAM_LANGUAGES,
        DIAGRAM_EXTENSIONS,
        IMAGE_EXTENSIONS,
    )
except (ImportError, ValueError):
    from .core.diagrams.bridge import (  # type: ignore
        calculate_aspect_dimensions,
        render_diagram_code_to_png,
        render_diagram_file_to_png,
        DIAGRAM_LANGUAGES,
        DIAGRAM_EXTENSIONS,
        IMAGE_EXTENSIONS,
    )

__all__ = [
    "calculate_aspect_dimensions",
    "render_diagram_code_to_png",
    "render_diagram_file_to_png",
    "DIAGRAM_LANGUAGES",
    "DIAGRAM_EXTENSIONS",
    "IMAGE_EXTENSIONS",
]
