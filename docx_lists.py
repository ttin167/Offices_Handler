"""
docx_lists.py — OpenXML Dynamic Numbering & List Tracker Facade.

Thin Facade pointing to core.docx.lists for strict Layered Modular Architecture.
"""
from __future__ import annotations

import os
import sys

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)

from core.docx.lists import *
import core.docx.lists as _mod

for _k in dir(_mod):
    if not _k.startswith("__"):
        globals()[_k] = getattr(_mod, _k)

