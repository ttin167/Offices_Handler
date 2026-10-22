"""
smart_post_processor.py — Universal Post-Processor & Ground Truth Healer Facade.

Thin Facade pointing to core.docx.post_processor for strict Layered Modular Architecture.
"""
from __future__ import annotations

import os
import sys

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)

from core.docx.post_processor import *
import core.docx.post_processor as _mod

for _k in dir(_mod):
    if not _k.startswith("__"):
        globals()[_k] = getattr(_mod, _k)

