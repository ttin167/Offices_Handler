"""
core.printing — Enterprise Physical & Virtual Printing Engine for Antigravity.

Components:
- PrinterManager: System printer discovery and silent print spooler.
- PrintOptimizer: Pre-print document optimizer (Grayscale, N-Up, Booklet, Watermark, Margin Guard).
"""

from __future__ import annotations

from .print_optimizer import PrintOptimizer, optimize_pdf_for_printing
from .printer_manager import PrinterManager

__all__ = [
    "PrinterManager",
    "PrintOptimizer",
    "optimize_pdf_for_printing",
]

