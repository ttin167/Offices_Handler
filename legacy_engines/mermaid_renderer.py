# -*- coding: utf-8 -*-
"""
legacy_engines/mermaid_renderer.py — Thin Backward-Compatibility Shim.
Redirects to core.diagrams.legacy.mermaid (Single Source of Truth).
"""
from core.diagrams.legacy.mermaid import *  # noqa: F401,F403
from core.diagrams.legacy.mermaid import (
    render_mermaid_to_png,
    calculate_aspect_dimensions,
    validate_mermaid_syntax,
    MERMAID_THEMES,
)

__all__ = [
    "render_mermaid_to_png",
    "calculate_aspect_dimensions",
    "validate_mermaid_syntax",
    "MERMAID_THEMES",
]
