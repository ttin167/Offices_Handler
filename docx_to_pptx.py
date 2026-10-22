"""
docx_to_pptx.py — Universal DOCX/PDF to 16:9 Presentation Synthesizer Facade.

Thin Facade pointing to core.pptx.docx_to_pptx for strict Layered Modular Architecture.
"""
from __future__ import annotations

import os
import sys

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)

from core.pptx.docx_to_pptx import *
import core.pptx.docx_to_pptx as _mod

for _k in dir(_mod):
    if not _k.startswith("__"):
        globals()[_k] = getattr(_mod, _k)

