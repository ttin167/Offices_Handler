"""
xlsx_writer.py — Universal Excel Mutation & Prototype Cloning Engine Facade.

Thin Facade pointing to core.xlsx.writer for strict Layered Modular Architecture.
"""
from __future__ import annotations

import os
import sys

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)

from core.xlsx.writer import *
import core.xlsx.writer as _mod
from core.xlsx.streaming import *
import core.xlsx.streaming as _stream_mod

for _k in dir(_mod):
    if not _k.startswith("__"):
        globals()[_k] = getattr(_mod, _k)

for _k in dir(_stream_mod):
    if not _k.startswith("__"):
        globals()[_k] = getattr(_stream_mod, _k)

