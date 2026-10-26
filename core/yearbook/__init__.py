"""
core/yearbook — Digital Yearbook & Pre-Press Photobook Studio Subsystem.

Provides:
- Intelligent layout engine with aspect-ratio clustering and 9 editorial spread templates.
- Industrial pre-press PDF export engine (300 DPI, 3mm Bleed, Crop Marks, Gutter compensation, Vietnamese Unicode font embedding).
- Designer handoff packager (.zip containing spread preview PDF, print-ready PDF, sorted high-res photos, Word DOCX brief, and project spec).
"""

from __future__ import annotations

from .layout_engine import (
    PRESETS,
    SPREAD_TEMPLATES,
    YearbookLayoutEngine,
    get_default_mock_spec,
)
from .pdf_press_exporter import (
    MM_TO_PT,
    YearbookPressEngine,
)
from .handoff_packager import (
    YearbookHandoffPackager,
)

__all__ = [
    "PRESETS",
    "SPREAD_TEMPLATES",
    "YearbookLayoutEngine",
    "get_default_mock_spec",
    "MM_TO_PT",
    "YearbookPressEngine",
    "YearbookHandoffPackager",
]
