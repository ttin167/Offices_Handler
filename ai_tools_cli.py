"""
ai_tools_cli.py — Master AI Tools CLI Dispatcher Facade.

Thin Facade pointing to cli.ai_tools for strict Layered Modular Architecture.
"""
from __future__ import annotations

import os
import sys

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)

from cli.ai_tools import *
import cli.ai_tools as _mod

for _k in dir(_mod):
    if not _k.startswith("__"):
        globals()[_k] = getattr(_mod, _k)


if __name__ == "__main__":
    from cli.ai_tools import main as _cli_main
    sys.exit(_cli_main())
