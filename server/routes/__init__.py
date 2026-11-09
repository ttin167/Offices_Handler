# -*- coding: utf-8 -*-
"""
server.routes — Modular REST API Route Handlers for Antigravity Universal Office Studio.
"""

from __future__ import annotations

from .base import Router, router, send_json, send_error, read_json_body

# Import route modules to auto-register decorators
from . import system_routes
from . import print_routes
from . import doc_routes
from . import copilot_routes

try:
    from . import ba_academy_routes  # noqa: F401
except ImportError:
    pass

__all__ = [
    "Router",
    "router",
    "send_json",
    "send_error",
    "read_json_body",
]

