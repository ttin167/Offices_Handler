# -*- coding: utf-8 -*-
"""
core/pipeline/ — Declarative Document Workflow Pipeline & Bulk Folder Runner (v4.6).
"""

from __future__ import annotations

from .engine import (
    WorkflowPipeline,
    get_default_max_workers,
    run_bulk_directory_pipeline,
    run_pipeline,
)

__all__ = [
    "WorkflowPipeline",
    "get_default_max_workers",
    "run_bulk_directory_pipeline",
    "run_pipeline",
]

