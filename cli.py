"""
cli.py — Legacy Document Converter CLI Facade.

Thin Facade pointing to cli.converter_cli for strict Layered Modular Architecture.
"""
from __future__ import annotations

import os
import sys

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)

from cli.converter_cli import *
import cli.converter_cli as _mod

for _k in dir(_mod):
    if not _k.startswith("__"):
        globals()[_k] = getattr(_mod, _k)


if __name__ == "__main__":
    from cli.converter_cli import run_cli as _cli_main
    sys.exit(_cli_main())
