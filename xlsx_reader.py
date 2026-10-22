"""
xlsx_reader.py — Universal Excel Sheet Reader & Token Extractor Facade.

Thin Facade pointing to core.xlsx.reader for strict Layered Modular Architecture.
"""
from __future__ import annotations

import os
import sys

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)

from core.xlsx.reader import *
import core.xlsx.reader as _mod

for _k in dir(_mod):
    if not _k.startswith("__"):
        globals()[_k] = getattr(_mod, _k)

