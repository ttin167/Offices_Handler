"""
job_manager.py — Root Facade for Background Job Queue & SSE Dispatcher.
"""

from __future__ import annotations

try:
    from core.job_manager import JobManager, global_job_manager
except (ImportError, ValueError):
    from .core.job_manager import JobManager, global_job_manager  # type: ignore

__all__ = ["JobManager", "global_job_manager"]
