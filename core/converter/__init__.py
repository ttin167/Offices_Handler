"""
core.converter — Universal multi-format conversion dispatchers (PDF ↔ DOCX ↔ MD ↔ PPTX).
"""

from __future__ import annotations

from .engine import convert_file, parse_page_range, get_pdf_page_count
from .markdown import (
    markdown_to_docx,
    markdown_to_pdf,
)

__all__ = [
    "convert_file",
    "parse_page_range",
    "get_pdf_page_count",
    "markdown_to_docx",
    "markdown_to_pdf",
]
