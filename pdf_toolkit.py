"""
pdf_toolkit.py — Root Facade for Everyday PDF Toolkit.

Thin Facade pointing to core.converter.pdf_toolkit for strict Layered Modular Architecture.
"""
from __future__ import annotations

from core.converter.pdf_toolkit import (
    merge_pdfs,
    split_pdf,
    rotate_and_clean_pdf,
    stamp_watermark_pdf,
    compress_pdf,
)

__all__ = [
    "merge_pdfs",
    "split_pdf",
    "rotate_and_clean_pdf",
    "stamp_watermark_pdf",
    "compress_pdf",
]
