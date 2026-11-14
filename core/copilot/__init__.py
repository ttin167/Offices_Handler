# -*- coding: utf-8 -*-
"""
core.copilot — Universal Multi-Format SQLite FTS5 RAG & Hybrid Office AI Copilot.
"""

from __future__ import annotations

from .copilot_provider import (
    SYSTEM_COPILOT_PROMPT,
    BaseCopilotProvider,
    CopilotActionCard,
    CopilotOrchestrator,
    CopilotResponse,
    GoogleGeminiProvider,
    LocalOllamaProvider,
    infer_action_cards,
    synthesize_offline_answer,
)
from .universal_rag import DEFAULT_RAG_DB_PATH, DocumentChunk, UniversalRAGEngine

__all__ = [
    "DEFAULT_RAG_DB_PATH",
    "DocumentChunk",
    "UniversalRAGEngine",
    "SYSTEM_COPILOT_PROMPT",
    "CopilotActionCard",
    "CopilotResponse",
    "BaseCopilotProvider",
    "LocalOllamaProvider",
    "GoogleGeminiProvider",
    "CopilotOrchestrator",
    "infer_action_cards",
    "synthesize_offline_answer",
]

