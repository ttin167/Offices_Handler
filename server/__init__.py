"""
server — High-performance zero-dependency Web API Bridge and persistence.
"""

from __future__ import annotations

from .app import run_server, PORT
from .db import (
    record_history,
    get_recent_history,
    delete_history_item,
    clear_history,
    init_db,
)

__all__ = [
    "run_server",
    "PORT",
    "record_history",
    "get_recent_history",
    "delete_history_item",
    "clear_history",
    "init_db",
]
