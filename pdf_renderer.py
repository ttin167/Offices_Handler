"""
pdf_renderer.py — Root Facade for Universal Cross-Platform PDF Renderer (v4.0).
Provides dual-fallback import to core/converter/pdf_renderer.py.
"""

from __future__ import annotations

import sys

try:
    from core.converter.pdf_renderer import (
        UniversalPdfRenderer,
        convert_document_to_pdf,
        find_libreoffice_binary,
        is_win32_com_available,
    )
except ImportError:
    import os
    sys.path.insert(0, os.path.dirname(__file__))
    from core.converter.pdf_renderer import (  # type: ignore
        UniversalPdfRenderer,
        convert_document_to_pdf,
        find_libreoffice_binary,
        is_win32_com_available,
    )

__all__ = [
    "UniversalPdfRenderer",
    "convert_document_to_pdf",
    "find_libreoffice_binary",
    "is_win32_com_available",
]
