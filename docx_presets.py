"""
docx_presets.py — Root Facade for Document Presets & Administrative Standards.

Thin Facade pointing to core.docx.presets for strict Layered Modular Architecture.
"""
from __future__ import annotations

from core.docx.presets import (
    list_available_presets,
    load_preset,
    apply_preset_to_document,
    build_administrative_header,
    BUILTIN_ND30_PRESET,
)

__all__ = [
    "list_available_presets",
    "load_preset",
    "apply_preset_to_document",
    "build_administrative_header",
    "BUILTIN_ND30_PRESET",
]
