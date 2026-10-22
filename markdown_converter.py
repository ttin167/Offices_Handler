"""
markdown_converter.py — Universal Decoupled Markdown Parser & Serializer Facade.

Thin Facade pointing to core.converter.markdown for strict Layered Modular Architecture.
"""
from __future__ import annotations

import os
import sys

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)

from core.converter.markdown import *
import core.converter.markdown as _mod

for _k in dir(_mod):
    if not _k.startswith("__"):
        globals()[_k] = getattr(_mod, _k)

