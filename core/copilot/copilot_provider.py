# -*- coding: utf-8 -*-
"""
core.copilot.copilot_provider — Polymorphic Hybrid LLM Connector for Antigravity Office Suite.

Supports two primary channels via BaseCopilotProvider:
1. LocalOllamaProvider (Default for air-gapped / enterprise privacy):
   Connects to OpenAI-compatible local endpoint (http://localhost:11434/v1/chat/completions)
   running Qwen 2.5 Coder / Llama 3.2.
2. GoogleGeminiProvider (High-Context Cloud Connector):
   Connects to Google Gemini API (1M-2M context window) for large PDF/Word/Excel analysis.
3. Deterministic RAG Fallback:
   When offline or when no LLM daemon/API key is active, synthesizes grounded responses
   and Action Cards directly from retrieved FTS5 context chunks.
"""

from __future__ import annotations

import abc
import json
import os
import time
import urllib.error
import urllib.request
from dataclasses import asdict, dataclass, field
from typing import Any, Dict, List, Optional


@dataclass
class CopilotActionCard:
    """Represents an executable UI Action Card suggested by the AI Copilot."""

    action_id: str
    title: str
    description: str
    command_type: str  # e.g. 'open_modal', 'switch_tab', 'open_viewer'
    payload: Dict[str, Any] = field(default_factory=dict)
    badge: str = "AI Action"

    def to_dict(self) -> Dict[str, Any]:
        return asdict(self)


@dataclass
class CopilotResponse:
    """Structured response returned by any Copilot provider."""

    success: bool
    provider: str
    model: str
    answer: str
    action_cards: List[CopilotActionCard] = field(default_factory=list)
    citations: List[Dict[str, Any]] = field(default_factory=list)
    latency_ms: float = 0.0
    fallback_used: bool = False
    error: Optional[str] = None

    def to_dict(self) -> Dict[str, Any]:
        return {
            "success": self.success,
            "provider": self.provider,
            "model": self.model,
            "answer": self.answer,
            "action_cards": [c.to_dict() for c in self.action_cards],
            "citations": self.citations,
            "latency_ms": round(self.latency_ms, 2),
            "fallback_used": self.fallback_used,
            "error": self.error,
        }


SYSTEM_COPILOT_PROMPT = (
    "Bạn là Antigravity AI Copilot — Trợ lý chuyên gia Tin học Văn phòng cho Antigravity Office Suite. "
    "Bạn hỗ trợ người dùng chuẩn hóa Word theo Nghị định 30/2020/NĐ-CP & trộn thư Mail Merge ({{SoTien:vn_words}}), "
    "xử lý dữ liệu Excel (làm sạch tên/SĐT/CCCD, khử trùng lặp, Pivot + biểu đồ, nối file XLOOKUP, 52+ công thức & AST Evaluator), "
    "xử lý PDF Pro (17 công cụ: ghép/tách, nén, đóng dấu giáp lai, ký số PAdES/PKCS#7, PDF/A-2b), "
    "tổng hợp PowerPoint 16:9 và chế bản in ấn công nghiệp (Vector Grayscale, Booklet Creep). "
    "Hãy trả lời ngắn gọn, chính xác, trích dẫn nguồn từ ngữ cảnh được cung cấp."
)


def infer_action_cards(prompt: str, context_chunks: Optional[List[Dict[str, Any]]] = None) -> List[CopilotActionCard]:
    """Infers contextual Web Studio Action Cards from pure Office IT user intent keywords."""
    q = (prompt or "").lower()
    cards: List[CopilotActionCard] = []

    if any(k in q for k in ("in ấn", "in tài liệu", "print", "booklet", "grayscale", "khay giấy")):
        cards.append(
            CopilotActionCard(
                action_id="act_print_studio",
                title="Mở Trạm In Ấn Công Nghiệp (Print Studio)",
                description="Cấu hình Vector Grayscale (/DeviceGray), bù trượt gáy sách Booklet Creep và chọn khay nạp giấy DEVMODE.",
                command_type="open_modal",
                payload={"target": "modal-print"},
                badge="Print Engine",
            )
        )

    if any(k in q for k in ("trộn thư", "mail merge", "ánh xạ", "hợp đồng hàng loạt", "vn_words", "đọc tiền")):
        cards.append(
            CopilotActionCard(
                action_id="act_mail_merge",
                title="Ánh Xạ Trường & Trộn Thư (Batch Mail Merge)",
                description="Nối biến {{token}}, {{SoTien:vn_words}} từ mẫu Word với dữ liệu Excel/CSV/JSON hàng loạt.",
                command_type="switch_tab",
                payload={"tab": "documents", "subtab": "mailmerge"},
                badge="Mail Merge",
            )
        )

    if any(k in q for k in ("nghị định 30", "nđ30", "văn bản hành chính", "quốc hiệu", "ghép word", "tách word", "watermark word", "số trang word")):
        cards.append(
            CopilotActionCard(
                action_id="act_word_studio",
                title="Mở Word Hành Chính NĐ 30/2020 & Tiện Ích",
                description="Chuẩn hóa thể thức Nghị định 30/2020/NĐ-CP, ghép/tách văn bản Word, đánh số trang X/Y và đóng dấu Watermark.",
                command_type="switch_tab",
                payload={"tab": "documents", "subtab": "wordstudio"},
                badge="Word NĐ30",
            )
        )

    if any(k in q for k in ("pdf", "giáp lai", "ký số", "pdf/a", "nén pdf", "ghép pdf", "tách pdf", "flatten")):
        cards.append(
            CopilotActionCard(
                action_id="act_pdf_pro",
                title="Mở PDF Pro Studio (17 Công Cụ)",
                description="Ghép/tách trang, nén PDF, đóng dấu giáp lai, ký số PKCS#7/PAdES, chuyển Ảnh ↔ PDF và chuẩn PDF/A-2b.",
                command_type="switch_tab",
                payload={"tab": "documents", "subtab": "pdftools"},
                badge="PDF Pro 17",
            )
        )

    if any(k in q for k in ("xlookup", "pivot", "làm sạch", "trùng lặp", "tách cột", "gộp cột", "nối file", "power tool")):
        cards.append(
            CopilotActionCard(
                action_id="act_xlsx_power",
                title="Mở Excel Power Data Suite & Pivot/XLOOKUP",
                description="Chuẩn hóa Tên/SĐT/CCCD VN, khử trùng lặp, tách/gộp cột, Pivot Chart tự động và nối file XLOOKUP không cần code.",
                command_type="switch_tab",
                payload={"tab": "spreadsheet", "subtab": "powertools"},
                badge="Excel Power",
            )
        )

    if any(k in q for k in ("theo dõi công việc", "daily report", "task_tracking", "project_tracking", "tiến độ")):
        cards.append(
            CopilotActionCard(
                action_id="act_xlsx_tracker",
                title="Mở Sổ Theo Dõi Công Việc & Báo Cáo Ngày",
                description="Quản lý danh mục đầu việc Excel và đồng bộ tự động Daily Report (project_tracking.xlsx / task_tracking.xlsx).",
                command_type="switch_tab",
                payload={"tab": "spreadsheet", "subtab": "tracker"},
                badge="Work Tracker",
            )
        )

    if any(k in q for k in ("công thức", "excel", "xlsx", "bảng tính", "sumif", "12 cổng", "ast")):
        cards.append(
            CopilotActionCard(
                action_id="act_spreadsheet_studio",
                title="Mở Spreadsheet Studio & 52+ Công Thức Excel",
                description="Kiểm định 12 cổng chất lượng, tính toán công thức AST ngoại tuyến và tra cứu 52+ công thức thực chiến.",
                command_type="switch_tab",
                payload={"tab": "spreadsheet", "subtab": "validator"},
                badge="Excel Studio",
            )
        )

    if any(k in q for k in ("xem trước", "preview", "viewer", "đọc file", "mở file")):
        first_file = None
        if context_chunks:
            for c in context_chunks:
                if c.get("filename"):
                    first_file = c.get("filename")
                    break
        cards.append(
            CopilotActionCard(
                action_id="act_doc_viewer",
                title="Mở Trình Xem Tài Liệu Trực Tiếp (In-Browser Viewer)",
                description="Xem trước DOCX, XLSX, PDF ngay trên trình duyệt với hiệu ứng G3 Shimmer Skeleton.",
                command_type="open_viewer",
                payload={"filename": first_file or ""},
                badge="Viewer G3",
            )
        )

    if not cards:
        cards.append(
            CopilotActionCard(
                action_id="act_qa_audit",
                title="Chạy Kiểm Toán Tài Liệu 13 Cổng (Unified QA Audit)",
                description="Kiểm tra toàn diện cấu trúc OpenXML, bảng biểu, công thức Excel và bảo mật tài liệu.",
                command_type="switch_tab",
                payload={"tab": "documents", "subtab": "qa"},
                badge="QA Audit",
            )
        )

    return cards


def synthesize_offline_answer(prompt: str, context_chunks: Optional[List[Dict[str, Any]]] = None) -> str:
    """Synthesizes a deterministic, grounded answer from FTS5 RAG chunks when LLM is offline."""
    chunks = context_chunks or []
    if not chunks:
        return (
            "Đã tiếp nhận yêu cầu: **\"" + prompt.strip() + "\"**.\n\n"
            "Hiện tại hệ thống đang hoạt động ở chế độ **Local-First Offline**. "
            "Bạn có thể nạp tài liệu Word, Excel hoặc PDF lên để lập chỉ mục vào **Universal RAG (SQLite FTS5)** "
            "hoặc kích hoạt trực tiếp các công cụ Tin học Văn phòng thông qua **Thẻ Hành Động (Action Cards)** bên dưới."
        )

    lines = [
        f"Dưới đây là kết quả phân tích từ **{len(chunks)} phân đoạn tài liệu** khớp nhất trong kho tri thức cục bộ (Universal FTS5 RAG):\n"
    ]
    for idx, ch in enumerate(chunks[:4], start=1):
        fname = ch.get("filename") or ch.get("doc_id") or f"Tài liệu #{idx}"
        sec = ch.get("section_title") or "Nội dung chính"
        snippet = (ch.get("snippet") or ch.get("content") or "").strip()
        if len(snippet) > 320:
            snippet = snippet[:320].rstrip() + "..."
        lines.append(f"{idx}. **[{fname} — {sec}]**: {snippet}")

    lines.append("\nBạn có thể chọn các **Thẻ Hành Động (Action Cards)** bên dưới để mở nhanh công cụ văn phòng tương ứng.")
    return "\n".join(lines)


class BaseCopilotProvider(abc.ABC):
    """Abstract base interface for all AI Copilot LLM connectors."""

    provider_name: str = "base"

    @abc.abstractmethod
    def generate_response(
        self,
        prompt: str,
        context_chunks: Optional[List[Dict[str, Any]]] = None,
        system_prompt: Optional[str] = None,
    ) -> CopilotResponse:
        """Generates a response with optional RAG context and Action Cards."""

    @abc.abstractmethod
    def health_check(self) -> Dict[str, Any]:
        """Checks connectivity and readiness of the underlying LLM endpoint."""


class LocalOllamaProvider(BaseCopilotProvider):
    """
    Local Offline LLM Provider connecting to an OpenAI-compatible local server (Ollama / vLLM).
    Ensures 100% document privacy on the workstation.
    """

    provider_name = "local_ollama"

    def __init__(
        self,
        base_url: Optional[str] = None,
        model: Optional[str] = None,
        timeout: float = 8.0,
    ) -> None:
        self.base_url = (
            base_url
            or os.environ.get("OLLAMA_BASE_URL")
            or "http://localhost:11434/v1"
        ).rstrip("/")
        self.model = model or os.environ.get("OLLAMA_MODEL") or "qwen2.5-coder:7b"
        self.timeout = timeout

    def health_check(self) -> Dict[str, Any]:
        models_url = f"{self.base_url}/models"
        req = urllib.request.Request(models_url, method="GET")
        try:
            with urllib.request.urlopen(req, timeout=1.5) as resp:
                status = resp.getcode()
                return {
                    "online": status == 200,
                    "provider": self.provider_name,
                    "base_url": self.base_url,
                    "model": self.model,
                }
        except Exception as exc:
            return {
                "online": False,
                "provider": self.provider_name,
                "base_url": self.base_url,
                "model": self.model,
                "fallback_ready": True,
                "detail": str(exc),
            }

    def generate_response(
        self,
        prompt: str,
        context_chunks: Optional[List[Dict[str, Any]]] = None,
        system_prompt: Optional[str] = None,
    ) -> CopilotResponse:
        t0 = time.perf_counter()
        chunks = context_chunks or []
        action_cards = infer_action_cards(prompt, chunks)
        citations = [
            {
                "doc_id": c.get("doc_id", ""),
                "filename": c.get("filename", ""),
                "section_title": c.get("section_title", ""),
            }
            for c in chunks[:5]
        ]

        context_text = "\n\n".join(
            f"[{c.get('filename', 'Doc')} - {c.get('section_title', '')}]: {c.get('content', '')}"
            for c in chunks[:5]
        )
        user_content = prompt if not context_text else f"Ngữ cảnh tài liệu:\n{context_text}\n\nCâu hỏi: {prompt}"

        payload = {
            "model": self.model,
            "messages": [
                {"role": "system", "content": system_prompt or SYSTEM_COPILOT_PROMPT},
                {"role": "user", "content": user_content},
            ],
            "temperature": 0.2,
            "stream": False,
        }

        chat_url = f"{self.base_url}/chat/completions"
        data_bytes = json.dumps(payload).encode("utf-8")
        req = urllib.request.Request(
            chat_url,
            data=data_bytes,
            headers={"Content-Type": "application/json"},
            method="POST",
        )

        try:
            with urllib.request.urlopen(req, timeout=self.timeout) as resp:
                raw = resp.read().decode("utf-8")
                parsed = json.loads(raw)
                choices = parsed.get("choices") or []
                answer = (
                    choices[0].get("message", {}).get("content", "")
                    if choices
                    else ""
                )
                if not answer:
                    answer = synthesize_offline_answer(prompt, chunks)
                elapsed = (time.perf_counter() - t0) * 1000.0
                return CopilotResponse(
                    success=True,
                    provider=self.provider_name,
                    model=self.model,
                    answer=answer,
                    action_cards=action_cards,
                    citations=citations,
                    latency_ms=elapsed,
                    fallback_used=False,
                )
        except Exception as exc:
            elapsed = (time.perf_counter() - t0) * 1000.0
            fallback_ans = synthesize_offline_answer(prompt, chunks)
            return CopilotResponse(
                success=True,
                provider=self.provider_name,
                model=f"{self.model} (Offline RAG Synthesis)",
                answer=fallback_ans,
                action_cards=action_cards,
                citations=citations,
                latency_ms=elapsed,
                fallback_used=True,
                error=str(exc),
            )


class GoogleGeminiProvider(BaseCopilotProvider):
    """
    Cloud High-Context LLM Provider using Google Gemini API (1M-2M token window)
    for analyzing large Word/Excel/PDF office documents.
    """

    provider_name = "cloud_gemini"

    def __init__(
        self,
        api_key: Optional[str] = None,
        model: Optional[str] = None,
        timeout: float = 15.0,
    ) -> None:
        self.api_key = (
            api_key
            or os.environ.get("GEMINI_API_KEY")
            or os.environ.get("GOOGLE_API_KEY")
            or ""
        ).strip()
        self.model = model or os.environ.get("GEMINI_MODEL") or "gemini-2.5-flash"
        self.timeout = timeout

    def health_check(self) -> Dict[str, Any]:
        has_key = bool(self.api_key)
        return {
            "online": has_key,
            "provider": self.provider_name,
            "model": self.model,
            "api_key_configured": has_key,
            "fallback_ready": True,
        }

    def generate_response(
        self,
        prompt: str,
        context_chunks: Optional[List[Dict[str, Any]]] = None,
        system_prompt: Optional[str] = None,
    ) -> CopilotResponse:
        t0 = time.perf_counter()
        chunks = context_chunks or []
        action_cards = infer_action_cards(prompt, chunks)
        citations = [
            {
                "doc_id": c.get("doc_id", ""),
                "filename": c.get("filename", ""),
                "section_title": c.get("section_title", ""),
            }
            for c in chunks[:8]
        ]

        if not self.api_key:
            elapsed = (time.perf_counter() - t0) * 1000.0
            return CopilotResponse(
                success=True,
                provider=self.provider_name,
                model=f"{self.model} (Offline RAG Synthesis)",
                answer=synthesize_offline_answer(prompt, chunks),
                action_cards=action_cards,
                citations=citations,
                latency_ms=elapsed,
                fallback_used=True,
                error="GEMINI_API_KEY is not configured; used local RAG synthesis.",
            )

        context_text = "\n\n".join(
            f"[{c.get('filename', 'Doc')} - {c.get('section_title', '')}]: {c.get('content', '')}"
            for c in chunks[:8]
        )
        full_prompt = (
            f"{system_prompt or SYSTEM_COPILOT_PROMPT}\n\n"
            f"Ngữ cảnh tài liệu:\n{context_text}\n\n"
            f"Yêu cầu của người dùng: {prompt}"
        )

        endpoint = (
            f"https://generativelanguage.googleapis.com/v1beta/models/"
            f"{self.model}:generateContent?key={self.api_key}"
        )
        payload = {
            "contents": [{"parts": [{"text": full_prompt}]}],
            "generationConfig": {"temperature": 0.2},
        }
        req = urllib.request.Request(
            endpoint,
            data=json.dumps(payload).encode("utf-8"),
            headers={"Content-Type": "application/json"},
            method="POST",
        )

        try:
            with urllib.request.urlopen(req, timeout=self.timeout) as resp:
                raw = resp.read().decode("utf-8")
                parsed = json.loads(raw)
                candidates = parsed.get("candidates") or []
                parts = (
                    candidates[0].get("content", {}).get("parts", [])
                    if candidates
                    else []
                )
                answer = parts[0].get("text", "") if parts else ""
                if not answer:
                    answer = synthesize_offline_answer(prompt, chunks)
                elapsed = (time.perf_counter() - t0) * 1000.0
                return CopilotResponse(
                    success=True,
                    provider=self.provider_name,
                    model=self.model,
                    answer=answer,
                    action_cards=action_cards,
                    citations=citations,
                    latency_ms=elapsed,
                    fallback_used=False,
                )
        except Exception as exc:
            elapsed = (time.perf_counter() - t0) * 1000.0
            return CopilotResponse(
                success=True,
                provider=self.provider_name,
                model=f"{self.model} (Offline RAG Synthesis)",
                answer=synthesize_offline_answer(prompt, chunks),
                action_cards=action_cards,
                citations=citations,
                latency_ms=elapsed,
                fallback_used=True,
                error=str(exc),
            )


class CopilotOrchestrator:
    """
    Polymorphic adapter selector supporting 'local', 'cloud', and 'hybrid' modes.
    """

    def __init__(
        self,
        default_mode: Optional[str] = None,
        local_provider: Optional[LocalOllamaProvider] = None,
        cloud_provider: Optional[GoogleGeminiProvider] = None,
    ) -> None:
        self.default_mode = (
            default_mode
            or os.environ.get("COPILOT_PROVIDER")
            or "local"
        ).lower()
        self.local_provider = local_provider or LocalOllamaProvider()
        self.cloud_provider = cloud_provider or GoogleGeminiProvider()

    def get_provider(self, mode: Optional[str] = None, context_length: int = 0) -> BaseCopilotProvider:
        target = (mode or self.default_mode).lower()
        if target in ("cloud", "gemini", "cloud_gemini"):
            return self.cloud_provider
        if target == "hybrid":
            # Route large contexts (> 16,000 chars) to Gemini Cloud if API key is present
            if context_length > 16000 and self.cloud_provider.api_key:
                return self.cloud_provider
            return self.local_provider
        return self.local_provider

    def chat(
        self,
        prompt: str,
        context_chunks: Optional[List[Dict[str, Any]]] = None,
        mode: Optional[str] = None,
        system_prompt: Optional[str] = None,
    ) -> CopilotResponse:
        chunks = context_chunks or []
        total_chars = sum(len(str(c.get("content", ""))) for c in chunks) + len(prompt or "")
        provider = self.get_provider(mode=mode, context_length=total_chars)
        return provider.generate_response(
            prompt=prompt,
            context_chunks=chunks,
            system_prompt=system_prompt,
        )

    def status(self) -> Dict[str, Any]:
        return {
            "default_mode": self.default_mode,
            "local": self.local_provider.health_check(),
            "cloud": self.cloud_provider.health_check(),
        }

