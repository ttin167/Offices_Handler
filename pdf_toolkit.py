"""
pdf_toolkit.py — Root Facade for Everyday PDF Toolkit.

Thin Facade pointing to core.converter.pdf_toolkit for strict Layered Modular Architecture.
"""
from __future__ import annotations

from core.converter.pdf_toolkit import (
    SENSITIVE_PATTERNS,
    compress_pdf,
    convert_images_to_pdf,
    export_pdf_to_images_zip,
    extract_pdf_tables_to_xlsx,
    flatten_and_sanitize_pdf,
    merge_pdfs,
    normalize_and_number_pdf,
    protect_pdf,
    redact_sensitive_pdf,
    render_pdf_thumbnails,
    reorganize_pdf_pages,
    rotate_and_clean_pdf,
    sign_pdf_with_stamp,
    split_pdf,
    stamp_giap_lai_pdf,
    stamp_watermark_pdf,
)

__all__ = [
    "SENSITIVE_PATTERNS",
    "merge_pdfs",
    "split_pdf",
    "rotate_and_clean_pdf",
    "stamp_watermark_pdf",
    "compress_pdf",
    "redact_sensitive_pdf",
    "protect_pdf",
    "normalize_and_number_pdf",
    "stamp_giap_lai_pdf",
    "extract_pdf_tables_to_xlsx",
    "render_pdf_thumbnails",
    "reorganize_pdf_pages",
    "convert_images_to_pdf",
    "export_pdf_to_images_zip",
    "sign_pdf_with_stamp",
    "flatten_and_sanitize_pdf",
]
