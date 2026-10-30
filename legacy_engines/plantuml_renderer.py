# -*- coding: utf-8 -*-
"""
legacy_engines/plantuml_renderer.py — Thin Backward-Compatibility Shim.
Redirects to core.diagrams.legacy.plantuml (Single Source of Truth).
"""
from core.diagrams.legacy.plantuml import *  # noqa: F401,F403
from core.diagrams.legacy.plantuml import (
    render_plantuml_to_png,
    calculate_aspect_dimensions,
    validate_plantuml_syntax,
)

__all__ = [
    "render_plantuml_to_png",
    "calculate_aspect_dimensions",
    "validate_plantuml_syntax",
]
