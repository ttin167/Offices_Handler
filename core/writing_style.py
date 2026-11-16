# -*- coding: utf-8 -*-
"""
core.writing_style — Enterprise Vietnamese & Technical Prose / Writing Style Engine (Văn Phong / Hành Văn).

Provides a clean, pure-core engine (no dependencies on `server/`, `ui/`, or `cli/`) for:
1. Standardized Writing Style Profiles (`WRITING_STYLE_REGISTRY`):
   - `hanh_chinh` (Nghị định 30/2020/NĐ-CP công vụ / hành chính)
   - `hoc_thuat` (Học thuật / Luận văn / Nghiên cứu khoa học)
   - `doanh_nghiep` (Doanh nghiệp / Báo cáo điều hành BLUF / Executive)
   - `phap_ly` (Pháp lý / Hợp đồng / Quy chế tuân thủ)
   - `ky_thuat` (Đặc tả kỹ thuật / SRS / Tài liệu kỹ thuật)
2. Deterministic Offline Prose Analysis (`analyze_writing_style`):
   - Sentence/word/paragraph readability metrics, transition density, suitability score (0–100),
     and rule-based findings (shorthand, colloquialisms, informal pronouns, wordiness, vagueness,
     duplicate words, and punctuation defects).
3. Safe Protected-Span Prose Refinement (`refine_writing_style`):
   - Protects `{{TOKEN}}`, inline/fenced code, URLs, and quoted spans (`"..."`, `“...”`) via masking.
   - Applies table-driven case-preserving deterministic replacements (`safe` or `strict` mode) +
     Vietnamese NFC/punctuation normalization.
4. Hybrid AI Rewriting (`rewrite_with_style`):
   - Builds style-grounded Copilot prompts and delegates to `CopilotOrchestrator` when `use_ai=True`,
     automatically falling back to deterministic `refine_writing_style` offline.
5. Administrative & Document Phrase Templates (`get_phrase_templates`):
   - Curated Vietnamese opening, basis, transition, request, and closing clauses by style & doc type.
6. Run-Safe Word `.docx` Integration (`apply_writing_style_to_docx`, `analyze_docx_writing_style`):
   - Mutates paragraph runs preserving OpenXML styling (`ERR_DOCX_002`), skipping drawings/fields/math.
"""

from __future__ import annotations

from dataclasses import asdict, dataclass, field
import os
from pathlib import Path
import re
from typing import Any, Callable, Dict, List, Optional, Sequence, Tuple
import unicodedata


# ---------------------------------------------------------------------------
# Data Models
# ---------------------------------------------------------------------------

@dataclass(frozen=True)
class StyleProfile:
    """Metadata and rhetorical rules defining a target prose register."""

    style_id: str
    name_vi: str
    name_en: str
    description: str
    aliases: Tuple[str, ...]
    tone_keywords: Tuple[str, ...]
    preferred_pronouns: Tuple[str, ...]
    forbidden_patterns_note: str
    target_avg_sentence_words: Tuple[int, int]
    max_sentence_words: int
    recommended_connectors: Tuple[str, ...]
    ai_system_directive: str

    def to_dict(self) -> Dict[str, Any]:
        return {
            "style_id": self.style_id,
            "name_vi": self.name_vi,
            "name_en": self.name_en,
            "description": self.description,
            "aliases": list(self.aliases),
            "tone_keywords": list(self.tone_keywords),
            "preferred_pronouns": list(self.preferred_pronouns),
            "forbidden_patterns_note": self.forbidden_patterns_note,
            "target_avg_sentence_words": list(self.target_avg_sentence_words),
            "max_sentence_words": self.max_sentence_words,
            "recommended_connectors": list(self.recommended_connectors),
        }


@dataclass(frozen=True)
class StyleRule:
    """Table-driven pattern rule for prose analysis and deterministic refinement."""

    rule_id: str
    category: str  # shorthand | colloquial | pronoun | wordiness | vagueness | redundancy
    pattern: str
    replacements: Dict[str, str]
    message: str
    severity: str = "warning"  # info | warning | error
    styles: Tuple[str, ...] = ()  # empty => applies to all styles
    min_mode: str = "safe"  # safe | strict (when auto-refining)
    case_sensitive: bool = False
    analyze_only: bool = False


@dataclass
class StyleFinding:
    """Single prose defect or suggestion identified during analysis."""

    rule_id: str
    category: str
    severity: str
    matched_text: str
    suggestion: str
    message: str
    paragraph_index: int = 0
    start: int = 0
    end: int = 0

    def to_dict(self) -> Dict[str, Any]:
        return asdict(self)


@dataclass
class StyleAnalysisResult:
    """Quantitative and diagnostic report of a document's writing style."""

    style_id: str
    style_name: str
    suitability_score: int
    metrics: Dict[str, Any]
    findings: List[StyleFinding] = field(default_factory=list)
    recommendations: List[str] = field(default_factory=list)

    def to_dict(self) -> Dict[str, Any]:
        return {
            "style_id": self.style_id,
            "style_name": self.style_name,
            "suitability_score": self.suitability_score,
            "metrics": self.metrics,
            "total_findings": len(self.findings),
            "findings": [f.to_dict() for f in self.findings],
            "recommendations": self.recommendations,
        }


@dataclass
class StyleRefinementResult:
    """Output bundle from deterministic or hybrid prose style transformation."""

    success: bool
    style_id: str
    style_name: str
    mode: str
    original_text: str
    refined_text: str
    changed: bool
    changes_count: int
    applied_edits: List[Dict[str, Any]] = field(default_factory=list)
    analysis_before: Optional[Dict[str, Any]] = None
    analysis_after: Optional[Dict[str, Any]] = None
    ai_used: bool = False
    fallback_used: bool = False
    provider: str = "deterministic_offline"
    error: Optional[str] = None

    def to_dict(self) -> Dict[str, Any]:
        return asdict(self)


# ---------------------------------------------------------------------------
# Style Registry (5 Core Vietnamese & Enterprise Prose Registers)
# ---------------------------------------------------------------------------

WRITING_STYLE_REGISTRY: Dict[str, StyleProfile] = {
    "hanh_chinh": StyleProfile(
        style_id="hanh_chinh",
        name_vi="Hành chính Công vụ (NĐ 30/2020)",
        name_en="Vietnamese Administrative / Official",
        description=(
            "Trang trọng, chuẩn mực, khách quan, đơn nghĩa theo Nghị định 30/2020/NĐ-CP "
            "về công tác văn thư (Công văn, Tờ trình, Quyết định, Báo cáo)."
        ),
        aliases=("nd30", "nd30_hanh_chinh", "administrative", "cong_vu", "official"),
        tone_keywords=("trang trọng", "chuẩn mực", "khách quan", "đơn nghĩa", "chính xác"),
        preferred_pronouns=("Cơ quan", "Đơn vị", "Quý Cơ quan", "Bộ phận chuyên môn"),
        forbidden_patterns_note="Cấm khẩu ngữ, viết tắt chat, từ cảm thán, đại từ thân mật (mình/bạn/nhé/ạ).",
        target_avg_sentence_words=(14, 28),
        max_sentence_words=40,
        recommended_connectors=(
            "Căn cứ",
            "Trên cơ sở đó",
            "Theo đó",
            "Tuy nhiên",
            "Do vậy",
            "Kính đề nghị",
            "Trân trọng",
        ),
        ai_system_directive=(
            "Hãy biên tập văn bản theo chuẩn văn phong Hành chính Công vụ Việt Nam (Nghị định 30/2020/NĐ-CP): "
            "trang trọng, khách quan, rõ chủ thể thực hiện, dùng từ chuẩn công vụ (Căn cứ, Trên cơ sở, Đề nghị, Thực hiện), "
            "loại bỏ hoàn toàn khẩu ngữ, cảm tính và viết tắt không chuẩn, giữ nguyên 100% số liệu và tên riêng."
        ),
    ),
    "hoc_thuat": StyleProfile(
        style_id="hoc_thuat",
        name_vi="Học thuật & Nghiên cứu",
        name_en="Academic & Research",
        description=(
            "Khách quan, phi nhân xưng, lập luận chặt chẽ có dẫn chứng logic cho luận văn, "
            "báo cáo khoa học và khóa luận tốt nghiệp."
        ),
        aliases=("academic", "academic_thesis", "thesis", "research", "luan_van"),
        tone_keywords=("khách quan", "phi nhân xưng", "lập luận", "chặt chẽ", "dẫn chứng"),
        preferred_pronouns=("Nghiên cứu này", "Tác giả", "Nhóm nghiên cứu", "Kết quả phân tích"),
        forbidden_patterns_note="Tránh ngôi thứ nhất thân mật (mình/em/tớ), khẳng định cảm tính cực đoan, khẩu ngữ.",
        target_avg_sentence_words=(16, 30),
        max_sentence_words=42,
        recommended_connectors=(
            "Tuy nhiên",
            "Mặt khác",
            "Bên cạnh đó",
            "Do đó",
            "Từ kết quả trên",
            "Nghiên cứu cho thấy",
            "Nhằm mục tiêu",
        ),
        ai_system_directive=(
            "Hãy biên tập văn bản theo văn phong Học thuật & Nghiên cứu khoa học: "
            "khách quan, trung tính, ưu tiên chủ ngữ phi nhân xưng (Nghiên cứu này, Kết quả phân tích), "
            "liên kết mạch lạc bằng từ nối logic, tránh cường điệu cảm tính, giữ nguyên toàn bộ thuật ngữ và trích dẫn."
        ),
    ),
    "doanh_nghiep": StyleProfile(
        style_id="doanh_nghiep",
        name_vi="Doanh nghiệp & Điều hành (BLUF)",
        name_en="Executive & Corporate",
        description=(
            "Súc tích, trực diện, hướng hành động và kết quả (Bottom Line Up Front), "
            "tối ưu cho báo cáo lãnh đạo, đề xuất dự án và email đối tác."
        ),
        aliases=("executive", "corporate", "corporate_report", "business", "dieu_hanh"),
        tone_keywords=("súc tích", "trực diện", "hướng kết quả", "chuyên nghiệp", "rõ trách nhiệm"),
        preferred_pronouns=("Đơn vị", "Đội ngũ", "Doanh nghiệp", "Quý Đối tác", "Ban Lãnh đạo"),
        forbidden_patterns_note="Tránh rào đón dài dòng, câu bị động mơ hồ, khẩu ngữ thiếu chuyên nghiệp.",
        target_avg_sentence_words=(12, 24),
        max_sentence_words=35,
        recommended_connectors=(
            "Trọng tâm",
            "Kết quả chính",
            "Theo đó",
            "Đề xuất",
            "Trên cơ sở đó",
            "Ưu tiên tiếp theo",
        ),
        ai_system_directive=(
            "Hãy biên tập văn bản theo văn phong Doanh nghiệp & Điều hành (BLUF — kết luận và số liệu lên đầu): "
            "ngắn gọn, súc tích, câu chủ động, rõ hành động và người chịu trách nhiệm, giữ nguyên mọi chỉ số KPI."
        ),
    ),
    "phap_ly": StyleProfile(
        style_id="phap_ly",
        name_vi="Pháp lý & Hợp đồng",
        name_en="Legal & Compliance",
        description=(
            "Chặt chẽ, tường minh quyền và nghĩa vụ, định lượng điều kiện và thời hạn, "
            "loại bỏ tuyệt đối từ ngữ đa nghĩa gây tranh chấp."
        ),
        aliases=("legal", "contract", "hop_dong", "compliance"),
        tone_keywords=("tường minh", "ràng buộc", "đơn nghĩa", "điều kiện", "thời hạn"),
        preferred_pronouns=("Các Bên", "Bên A", "Bên B", "Bên Cung cấp", "Bên Sử dụng"),
        forbidden_patterns_note="Cảnh báo từ ngữ mơ hồ (sớm nhất có thể, hợp lý tùy tình hình, đại khái, khoảng).",
        target_avg_sentence_words=(16, 32),
        max_sentence_words=48,
        recommended_connectors=(
            "Trừ trường hợp",
            "Kể từ ngày",
            "Trong thời hạn",
            "Theo quy định tại",
            "Trên cơ sở thỏa thuận",
            "Các Bên thống nhất",
        ),
        ai_system_directive=(
            "Hãy biên tập văn bản theo văn phong Pháp lý & Hợp đồng: "
            "chặt chẽ, đơn nghĩa, phân định rõ chủ thể (Bên A/Bên B), nghĩa vụ, ngoại lệ và mốc thời hạn, "
            "loại bỏ từ ngữ ước lệ hoặc mơ hồ."
        ),
    ),
    "ky_thuat": StyleProfile(
        style_id="ky_thuat",
        name_vi="Đặc tả Kỹ thuật (SRS / Docs)",
        name_en="Technical Specification",
        description=(
            "Chính xác, nhất quán thuật ngữ, mô tả hành vi tất định (Input → Điều kiện → Output), "
            "phù hợp cho tài liệu kiến trúc, SRS, API và hướng dẫn vận hành."
        ),
        aliases=("technical", "spec", "srs", "engineering", "docs"),
        tone_keywords=("chính xác", "nhất quán", "tất định", "kiểm chứng được", "rõ ràng"),
        preferred_pronouns=("Hệ thống", "Mô-đun", "Người dùng", "Quản trị viên"),
        forbidden_patterns_note="Tránh tính từ mơ hồ không có ngưỡng đo (nhanh, tối ưu, thân thiện) và câu kể dài dòng.",
        target_avg_sentence_words=(12, 24),
        max_sentence_words=34,
        recommended_connectors=(
            "Khi",
            "Nếu",
            "Trong trường hợp",
            "Bắt buộc",
            "Mặc định",
            "Cụ thể",
        ),
        ai_system_directive=(
            "Hãy biên tập văn bản theo văn phong Đặc tả Kỹ thuật (SRS / Engineering Docs): "
            "chính xác, tất định, câu rõ điều kiện và kết quả đầu ra, bảo toàn 100% tên hàm, mã lỗi, tham số và code."
        ),
    ),
}

_STYLE_ALIAS_MAP: Dict[str, str] = {}
for _sid, _prof in WRITING_STYLE_REGISTRY.items():
    _STYLE_ALIAS_MAP[_sid] = _sid
    for _alias in _prof.aliases:
        _STYLE_ALIAS_MAP[_alias.lower()] = _sid


def resolve_style_id(style_id: Optional[str], default: str = "hanh_chinh") -> str:
    """Resolve user-provided style alias to canonical style ID."""
    if not style_id or not str(style_id).strip():
        return default
    key = str(style_id).strip().lower().replace("-", "_").replace(" ", "_")
    if key in _STYLE_ALIAS_MAP:
        return _STYLE_ALIAS_MAP[key]
    raise ValueError(
        f"Không hỗ trợ văn phong '{style_id}'. Các giá trị hợp lệ: {', '.join(WRITING_STYLE_REGISTRY.keys())}."
    )


def get_writing_style_profile(style_id: str = "hanh_chinh") -> StyleProfile:
    """Return StyleProfile dataclass for the canonical or aliased style ID."""
    canonical = resolve_style_id(style_id)
    return WRITING_STYLE_REGISTRY[canonical]


def list_writing_styles() -> List[Dict[str, Any]]:
    """Return metadata for all available writing style registers."""
    return [prof.to_dict() for prof in WRITING_STYLE_REGISTRY.values()]


# ---------------------------------------------------------------------------
# Rule Catalog (Table-Driven Prose Rules)
# ---------------------------------------------------------------------------

_DEFAULT_OK_MAP = {
    "hanh_chinh": "nhất trí",
    "hoc_thuat": "đạt yêu cầu",
    "doanh_nghiep": "thống nhất",
    "phap_ly": "chấp thuận",
    "ky_thuat": "hợp lệ",
}

_DEFAULT_EXTREME_MAP = {
    "hanh_chinh": "đặc biệt",
    "hoc_thuat": "đáng kể",
    "doanh_nghiep": "rõ rệt",
    "phap_ly": "nghiêm trọng",
    "ky_thuat": "đáng kể",
}

_DEFAULT_PRONOUN_1P_MAP = {
    "hanh_chinh": "đơn vị",
    "hoc_thuat": "nghiên cứu này",
    "doanh_nghiep": "đội ngũ",
    "phap_ly": "Bên đề xuất",
    "ky_thuat": "hệ thống",
}

_DEFAULT_PRONOUN_2P_MAP = {
    "hanh_chinh": "Quý Cơ quan",
    "hoc_thuat": "người đọc",
    "doanh_nghiep": "Quý Đối tác",
    "phap_ly": "Bên còn lại",
    "ky_thuat": "người dùng",
}

STYLE_RULES: Tuple[StyleRule, ...] = (
    # 1. Chat Shorthand & Unstandardized Abbreviations
    StyleRule(
        rule_id="STYLE_SHORTHAND_KO",
        category="shorthand",
        pattern=r"ko",
        replacements={"*": "không"},
        message="Tránh dùng chữ viết tắt chat 'ko' trong văn bản chuẩn.",
        severity="error",
        case_sensitive=True,
    ),
    StyleRule(
        rule_id="STYLE_SHORTHAND_K_PHRASE",
        category="shorthand",
        pattern=r"k(?=\s+(?:có|thể|cần|biết|phải|làm|được|đạt|nên|rõ)(?![\w\-/]))",
        replacements={"*": "không"},
        message="Thay viết tắt 'k' bằng 'không'.",
        severity="error",
        case_sensitive=True,
    ),
    StyleRule(
        rule_id="STYLE_SHORTHAND_DC_ACCENT",
        category="shorthand",
        pattern=r"đc",
        replacements={"*": "được"},
        message="Thay chữ viết tắt 'đc' bằng 'được'.",
        severity="error",
        case_sensitive=True,
    ),
    StyleRule(
        rule_id="STYLE_SHORTHAND_DC_PLAIN",
        category="shorthand",
        pattern=r"dc",
        replacements={"*": "được"},
        message="Thay chữ viết tắt 'dc' bằng 'được'.",
        severity="error",
        case_sensitive=True,
    ),
    StyleRule(
        rule_id="STYLE_SHORTHAND_NTN",
        category="shorthand",
        pattern=r"ntn",
        replacements={"*": "như thế nào"},
        message="Thay viết tắt 'ntn' bằng 'như thế nào'.",
        severity="error",
        case_sensitive=True,
    ),
    StyleRule(
        rule_id="STYLE_SHORTHAND_MN",
        category="shorthand",
        pattern=r"mn",
        replacements={
            "hanh_chinh": "các đơn vị liên quan",
            "hoc_thuat": "các đối tượng khảo sát",
            "doanh_nghiep": "các bộ phận liên quan",
            "phap_ly": "các bên liên quan",
            "ky_thuat": "người dùng",
        },
        message="Tránh viết tắt 'mn' trong văn bản chính thức.",
        severity="error",
        case_sensitive=True,
    ),
    StyleRule(
        rule_id="STYLE_SHORTHAND_SO_VS",
        category="shorthand",
        pattern=r"so\s+vs",
        replacements={"*": "so với"},
        message="Thay 'so vs' bằng 'so với' trong tiếng Việt.",
        severity="warning",
        case_sensitive=True,
    ),
    StyleRule(
        rule_id="STYLE_SHORTHAND_CUNG_VS",
        category="shorthand",
        pattern=r"cùng\s+vs",
        replacements={"*": "cùng với"},
        message="Thay 'cùng vs' bằng 'cùng với' trong tiếng Việt.",
        severity="warning",
        case_sensitive=True,
    ),
    StyleRule(
        rule_id="STYLE_SHORTHAND_TL",
        category="shorthand",
        pattern=r"tl",
        replacements={"*": "trả lời"},
        message="Tránh viết tắt 'tl' thay cho 'trả lời'.",
        severity="warning",
        case_sensitive=True,
        min_mode="strict",
    ),
    # 2. Colloquial & Spoken Particles
    StyleRule(
        rule_id="STYLE_COLLOQUIAL_RAT_OK",
        category="colloquial",
        pattern=r"(?:rất\s+là|rất|khá)\s+(?:okie|oke|ok)",
        replacements={
            "hanh_chinh": "đạt yêu cầu",
            "hoc_thuat": "đạt yêu cầu",
            "doanh_nghiep": "khả quan",
            "phap_ly": "phù hợp quy định",
            "ky_thuat": "đạt tiêu chuẩn",
        },
        message="Thay cụm khẩu ngữ 'rất/khá ok' bằng diễn đạt chuẩn mực.",
        severity="warning",
        case_sensitive=False,
    ),
    StyleRule(
        rule_id="STYLE_COLLOQUIAL_OK",
        category="colloquial",
        pattern=r"(?:okie|oke|ok)",
        replacements=_DEFAULT_OK_MAP,
        message="Thay từ khẩu ngữ 'ok/oke' bằng thuật ngữ chuẩn mực.",
        severity="warning",
        case_sensitive=False,
    ),
    StyleRule(
        rule_id="STYLE_COLLOQUIAL_PARTICLES",
        category="colloquial",
        pattern=r"(?:nhé|nha|nè|hả|vậy á|đấy ạ|đó ạ|ạ)(?=\s*(?:[,.;!?]|$))",
        replacements={"*": ""},
        message="Loại bỏ tiểu từ khẩu ngữ ('nhé', 'nha', 'nè', 'ạ') khỏi văn bản chính thức.",
        severity="warning",
        case_sensitive=False,
    ),
    StyleRule(
        rule_id="STYLE_COLLOQUIAL_RAT_LA",
        category="colloquial",
        pattern=r"rất\s+là",
        replacements={"*": "rất"},
        message="Rút gọn cụm khẩu ngữ 'rất là' thành 'rất'.",
        severity="info",
    ),
    StyleRule(
        rule_id="STYLE_COLLOQUIAL_EXTREME",
        category="colloquial",
        pattern=r"(?:cực\s+kỳ|siêu\s+tốt|siêu\s+nhanh|quá\s+trời|vô\s+cùng\s+tuyệt\s+vời)",
        replacements=_DEFAULT_EXTREME_MAP,
        message="Thay tính từ cảm tính cường điệu bằng diễn đạt khách quan, có kiểm chứng.",
        severity="warning",
    ),
    StyleRule(
        rule_id="STYLE_COLLOQUIAL_SURE_100",
        category="colloquial",
        pattern=r"chắc\s+chắn\s+100%",
        replacements={
            "hanh_chinh": "bảo đảm đầy đủ",
            "hoc_thuat": "có độ tin cậy cao",
            "doanh_nghiep": "cam kết đạt mức tối đa",
            "phap_ly": "bảo đảm tuân thủ tuyệt đối",
            "ky_thuat": "đáp ứng tất định",
        },
        message="Tránh lối nói tuyệt đối hóa mang tính khẩu ngữ ('chắc chắn 100%').",
        severity="warning",
    ),
    StyleRule(
        rule_id="STYLE_COLLOQUIAL_XONG_ROI",
        category="colloquial",
        pattern=r"(?:đã\s+xong\s+rồi|làm\s+xong\s+rồi|xong\s+xuôi)",
        replacements={"*": "đã hoàn tất"},
        message="Dùng 'đã hoàn tất' thay cho khẩu ngữ 'xong rồi/xong xuôi'.",
        severity="warning",
    ),
    StyleRule(
        rule_id="STYLE_COLLOQUIAL_LAM_CAI_NAY",
        category="colloquial",
        pattern=r"làm\s+cái\s+này",
        replacements={
            "hanh_chinh": "thực hiện nội dung này",
            "hoc_thuat": "triển khai nội dung này",
            "doanh_nghiep": "thực hiện hạng mục này",
            "phap_ly": "thực hiện nghĩa vụ nêu trên",
            "ky_thuat": "triển khai thành phần này",
        },
        message="Thay cụm khẩu ngữ 'làm cái này' bằng động từ chuyên ngành phù hợp.",
        severity="warning",
    ),
    StyleRule(
        rule_id="STYLE_COLLOQUIAL_CAI_NAY",
        category="colloquial",
        pattern=r"(?:mấy\s+cái\s+này|cái\s+này)",
        replacements={
            "hanh_chinh": "nội dung này",
            "hoc_thuat": "vấn đề này",
            "doanh_nghiep": "hạng mục này",
            "phap_ly": "nội dung nêu trên",
            "ky_thuat": "thành phần này",
        },
        message="Thay đại từ chỉ định khẩu ngữ ('cái này') bằng danh từ chính xác.",
        severity="warning",
    ),
    # 3. Wordiness & Filler Clauses
    StyleRule(
        rule_id="STYLE_WORDINESS_TIME",
        category="wordiness",
        pattern=r"trong\s+thời\s+gian\s+vừa\s+qua",
        replacements={"*": "trong thời gian qua"},
        message="Rút gọn 'trong thời gian vừa qua' thành 'trong thời gian qua'.",
        severity="info",
    ),
    StyleRule(
        rule_id="STYLE_WORDINESS_PRESENT",
        category="wordiness",
        pattern=r"vào\s+thời\s+điểm\s+hiện\s+tại",
        replacements={"*": "hiện nay"},
        message="Rút gọn 'vào thời điểm hiện tại' thành 'hiện nay'.",
        severity="info",
    ),
    StyleRule(
        rule_id="STYLE_WORDINESS_EXEC",
        category="wordiness",
        pattern=r"(?:tiến\s+hành\s+thực\s+hiện\s+việc|tiến\s+hành\s+thực\s+hiện|thực\s+hiện\s+việc)",
        replacements={"*": "thực hiện"},
        message="Loại bỏ động từ sáo rỗng lặp nghĩa ('tiến hành thực hiện việc').",
        severity="warning",
    ),
    StyleRule(
        rule_id="STYLE_WORDINESS_PURPOSE",
        category="wordiness",
        pattern=r"nhằm\s+mục\s+đích\s+để",
        replacements={"*": "nhằm"},
        message="Rút gọn cụm dư thừa 'nhằm mục đích để' thành 'nhằm'.",
        severity="info",
    ),
    StyleRule(
        rule_id="STYLE_WORDINESS_GENERALLY",
        category="wordiness",
        pattern=r"(?:nói\s+chung\s+là|nói\s+tóm\s+lại\s+là)",
        replacements={"*": "nhìn chung,"},
        message="Thay cụm khẩu ngữ chuyển ý bằng 'nhìn chung,'.",
        severity="info",
    ),
    StyleRule(
        rule_id="STYLE_WORDINESS_BASICALLY",
        category="wordiness",
        pattern=r"về\s+cơ\s+bản\s+thì",
        replacements={"*": "về cơ bản,"},
        message="Loại bỏ hư từ 'thì' trong 'về cơ bản thì'.",
        severity="info",
    ),
    StyleRule(
        rule_id="STYLE_WORDINESS_SEEN",
        category="wordiness",
        pattern=r"có\s+thể\s+thấy\s+rằng\s+một\s+điều\s+là",
        replacements={"*": "có thể nhận thấy"},
        message="Rút gọn mệnh đề thừa 'có thể thấy rằng một điều là'.",
        severity="info",
    ),
    # 4. Register-Inappropriate Informal Pronouns (Strict mode auto-replace; Always flagged)
    StyleRule(
        rule_id="STYLE_PRONOUN_1P",
        category="pronoun",
        pattern=r"(?:tụi\s+mình|bọn\s+mình|chúng\s+mình|(?<!tự\s)(?<!chính\s)mình|tớ)",
        replacements=_DEFAULT_PRONOUN_1P_MAP,
        message="Đại từ ngôi thứ nhất thân mật không phù hợp văn bản chính thức/học thuật/pháp lý.",
        severity="warning",
        min_mode="strict",
        case_sensitive=False,
    ),
    StyleRule(
        rule_id="STYLE_PRONOUN_2P",
        category="pronoun",
        pattern=r"(?:các\s+bạn|cậu)",
        replacements=_DEFAULT_PRONOUN_2P_MAP,
        message="Đại từ xưng hô thân mật không phù hợp ngữ vực văn bản.",
        severity="warning",
        min_mode="strict",
        case_sensitive=False,
    ),
    # 5. Vagueness / Ambiguity (Flagged especially for Legal & Technical specs)
    StyleRule(
        rule_id="STYLE_VAGUE_SOONEST",
        category="vagueness",
        pattern=r"(?:sớm\s+nhất\s+có\s+thể|càng\s+sớm\s+càng\s+tốt)",
        replacements={
            "phap_ly": "trong thời hạn ... ngày làm việc kể từ ngày ...",
            "ky_thuat": "trong ngưỡng SLA tối đa ... giây/giờ",
            "doanh_nghiep": "trước thời hạn đã thống nhất",
            "hanh_chinh": "theo đúng tiến độ quy định",
            "hoc_thuat": "trong khung thời gian xác định",
        },
        message="Cụm thời gian ước lệ thiếu mốc định lượng cụ thể.",
        severity="warning",
        styles=("phap_ly", "ky_thuat", "doanh_nghiep", "hanh_chinh"),
        analyze_only=True,
    ),
    StyleRule(
        rule_id="STYLE_VAGUE_REASONABLE",
        category="vagueness",
        pattern=r"(?:hợp\s+lý\s+tùy\s+tình\s+hình|tùy\s+tình\s+hình\s+thực\s+tế|đại\s+khái|tương\s+đối\s+ổn)",
        replacements={
            "phap_ly": "căn cứ tiêu chí định lượng được quy định tại Phụ lục ...",
            "ky_thuat": "theo ngưỡng cấu hình xác định",
            "*": "cần bổ sung tiêu chí đánh giá định lượng cụ thể",
        },
        message="Diễn đạt mơ hồ, dễ phát sinh cách hiểu khác nhau khi nghiệm thu hoặc tranh chấp.",
        severity="warning",
        styles=("phap_ly", "ky_thuat", "hoc_thuat"),
        analyze_only=True,
    ),
    StyleRule(
        rule_id="STYLE_VAGUE_TECH_ADJ",
        category="vagueness",
        pattern=r"(?:xử\s+lý\s+nhanh\s+chóng|giao\s+diện\s+thân\s+thiện|tối\s+ưu\s+hóa\s+tốt)",
        replacements={
            "ky_thuat": "đáp ứng chỉ số định lượng (ví dụ: độ trễ <= ... ms / đạt WCAG AA)",
            "*": "nêu rõ chỉ số đo lường cụ thể",
        },
        message="Tính từ mô tả phi định lượng trong tài liệu kỹ thuật.",
        severity="info",
        styles=("ky_thuat",),
        analyze_only=True,
    ),
)

_DUP_WORD_REGEX = re.compile(
    r"(?<![\w\-/])(và|các|những|của|trong|với|được|đã|sẽ|cho|về|tại|theo|khi|nếu)\s+\1(?![\w\-/])",
    re.IGNORECASE,
)


# ---------------------------------------------------------------------------
# Phrase Templates Catalog
# ---------------------------------------------------------------------------

PHRASE_TEMPLATES: Dict[str, Dict[str, Dict[str, Any]]] = {
    "hanh_chinh": {
        "cong_van": {
            "title": "Công văn hành chính (NĐ 30/2020/NĐ-CP)",
            "opening": "Thực hiện chỉ đạo của {{CapTren}} về việc {{NoiDung}}, {{DonVi}} trân trọng báo cáo/đề nghị nội dung sau:",
            "basis": "Căn cứ {{VanBanCanCu}} ngày {{NgayBanHanh}} của {{CoQuanBanHanh}} về {{TrichYeu}};",
            "transition": "Trên cơ sở rà soát thực tế và đối chiếu quy định hiện hành, {{DonVi}} kính đề nghị {{QuýCoQuan}} xem xét:",
            "closing": "Trân trọng cảm ơn sự phối hợp của {{QuýCoQuan}}./.",
        },
        "to_trinh": {
            "title": "Tờ trình phê duyệt",
            "opening": "Kính trình {{CapPheDuyet}} xem xét, phê duyệt {{TenDeXuat}} với các nội dung trọng tâm dưới đây:",
            "basis": "Căn cứ chức năng, nhiệm vụ được giao và yêu cầu thực tiễn công tác năm {{Nam}};",
            "transition": "Từ các cơ sở nêu trên, {{DonVi}} kính trình {{CapPheDuyet}} chấp thuận chủ trương triển khai phương án {{PhuongAn}}.",
            "closing": "Kính trình {{CapPheDuyet}} xem xét, quyết định./.",
        },
        "bao_cao": {
            "title": "Báo cáo tổng kết / tiến độ",
            "opening": "Thực hiện kế hoạch công tác {{KyBaoCao}}, {{DonVi}} báo cáo kết quả triển khai các nhiệm vụ trọng tâm như sau:",
            "basis": "So với mục tiêu đề ra tại Kế hoạch số {{SoKeHoach}}, tiến độ chung đạt {{TyLe}}%.",
            "transition": "Bên cạnh kết quả đạt được, báo cáo làm rõ những tồn tại, nguyên nhân và đề xuất giải pháp khắc phục trong kỳ tới.",
            "closing": "Trên đây là báo cáo kết quả thực hiện của {{DonVi}}, kính báo cáo {{CapTren}} theo dõi, chỉ đạo./.",
        },
        "quyet_dinh": {
            "title": "Quyết định cá biệt / quy phạm",
            "opening": "Căn cứ Quy chế tổ chức và hoạt động của {{CoQuan}}; xét đề nghị của {{TruongDonViDeNghi}},",
            "basis": "Điều 1. Ban hành/Phê duyệt kèm theo Quyết định này {{TenTaiLieu}}.",
            "transition": "Điều 2. Quyết định này có hiệu lực thi hành kể từ ngày ký.",
            "closing": "Điều 3. Chánh Văn phòng, Thủ trưởng các đơn vị liên quan và các cá nhân có tên chịu trách nhiệm thi hành Quyết định này./.",
        },
        "bien_ban": {
            "title": "Biên bản cuộc họp / nghiệm thu",
            "opening": "Hôm nay, vào lúc {{Gio}} giờ {{Phut}}, ngày {{Ngay}} tháng {{Thang}} năm {{Nam}}, tại {{DiaDiem}}, chúng tôi gồm:",
            "basis": "Cuộc họp tiến hành thảo luận và thống nhất các nội dung chi tiết dưới đây:",
            "transition": "Các thành viên tham dự nhất trí thông qua toàn bộ nội dung kết luận và phân công trách nhiệm thực hiện.",
            "closing": "Biên bản kết thúc vào lúc {{GioKetThuc}} cùng ngày, đã được thông qua và nhất trí ký tên xác nhận./.",
        },
    },
    "hoc_thuat": {
        "abstract": {
            "title": "Tóm tắt nghiên cứu (Abstract)",
            "opening": "Nghiên cứu này tập trung phân tích {{VanDeNghienCuu}} trong bối cảnh {{BoiCanh}} nhằm giải quyết hạn chế về {{KhoangTrong}}.",
            "basis": "Bằng phương pháp {{PhuongPhap}} trên tập dữ liệu {{DuLieu}}, nghiên cứu đã kiểm chứng giả thuyết đặt ra.",
            "transition": "Kết quả thực nghiệm cho thấy phương án đề xuất cải thiện {{ChiSo}} ở mức {{GiaTri}}, qua đó khẳng định tính khả thi khi áp dụng thực tiễn.",
            "closing": "Từ khóa: {{TuKhoa1}}; {{TuKhoa2}}; {{TuKhoa3}}.",
        },
        "conclusion": {
            "title": "Kết luận & Hướng phát triển",
            "opening": "Tổng hợp các kết quả lý thuyết và thực nghiệm cho thấy mục tiêu ban đầu của nghiên cứu đã được hoàn thành đầy đủ.",
            "basis": "Cụ thể, hệ thống/mô hình đạt độ chính xác {{DoChinhXac}} và duy trì tính ổn định qua các kịch bản kiểm thử.",
            "transition": "Tuy nhiên, nghiên cứu vẫn còn giới hạn nhất định về {{GioiHan}} cần được tiếp tục khảo sát ở quy mô rộng hơn.",
            "closing": "Trong giai đoạn tiếp theo, nhóm nghiên cứu định hướng mở rộng sang {{HuongMoRong}} nhằm nâng cao giá trị ứng dụng.",
        },
    },
    "doanh_nghiep": {
        "executive_brief": {
            "title": "Báo cáo nhanh Điều hành (BLUF)",
            "opening": "Kết luận trọng tâm: {{KetLuanChinh}} — đạt {{TyLeDat}}% kế hoạch với tác động tài chính ước tính {{GiaTri}}.",
            "basis": "Số liệu then chốt cho thấy {{ChiSo1}} tăng {{BienDong1}}, trong khi {{ChiSo2}} cần kiểm soát ở ngưỡng {{Nguong}}.",
            "transition": "Kiến nghị hành động ngay trong tuần tới: (1) {{HanhDong1}}; (2) {{HanhDong2}} do {{NguoiPhuTrach}} chủ trì.",
            "closing": "Kính trình Ban Lãnh đạo phê duyệt phương án ưu tiên để kịp tiến độ {{MocThoiGian}}.",
        },
        "client_email": {
            "title": "Thư trao đổi Đối tác / Khách hàng",
            "opening": "Kính gửi Quý {{TenDoiTac}}, cảm ơn Quý đơn vị đã phối hợp chặt chẽ trong hạng mục {{HangMuc}}.",
            "basis": "Chúng tôi xin gửi kèm tài liệu cập nhật phiên bản {{PhienBan}} cùng tóm tắt các điểm đã thống nhất.",
            "transition": "Để bảo đảm mốc bàn giao ngày {{NgayBanGiao}}, kính mong Quý đơn vị phản hồi xác nhận trước {{HanPhanHoi}}.",
            "closing": "Trân trọng, {{NguoiGui}} — {{ChucVu}}.",
        },
    },
    "phap_ly": {
        "obligation": {
            "title": "Điều khoản Nghĩa vụ các Bên",
            "opening": "Điều {{SoDieu}}. Quyền và nghĩa vụ của {{ChuThe}}",
            "basis": "{{ChuThe}} có nghĩa vụ thực hiện đầy đủ, đúng hạn {{NoiDungNghiaVu}} trong thời hạn {{SoNgay}} ngày làm việc kể từ ngày {{MocKichHoat}}.",
            "transition": "Trường hợp chậm trễ không do Sự kiện Bất khả kháng, bên vi phạm chịu trách nhiệm khắc phục trong vòng {{HanKhacPhuc}} ngày và bồi thường thiệt hại thực tế phát sinh.",
            "closing": "Mọi sửa đổi, bổ sung đối với điều khoản này chỉ có hiệu lực khi được đại diện có thẩm quyền của Các Bên ký xác nhận bằng văn bản.",
        },
        "force_majeure": {
            "title": "Điều khoản Bất khả kháng",
            "opening": "Sự kiện Bất khả kháng là sự kiện xảy ra một cách khách quan, không thể lường trước được và không thể khắc phục được mặc dù đã áp dụng mọi biện pháp cần thiết.",
            "basis": "Bên bị ảnh hưởng phải thông báo bằng văn bản cho Bên còn lại trong thời hạn {{SoNgayThongBao}} ngày làm việc kể từ thời điểm phát sinh sự kiện.",
            "transition": "Thời gian thực hiện nghĩa vụ bị gián đoạn được gia hạn tương ứng với thời gian kéo dài của Sự kiện Bất khả kháng.",
            "closing": "Nếu Sự kiện Bất khả kháng kéo dài quá {{SoNgayToiDa}} ngày liên tục, mỗi Bên có quyền đề nghị đàm phán chấm dứt hợp đồng theo quy định.",
        },
    },
    "ky_thuat": {
        "requirement": {
            "title": "Đặc tả Yêu cầu Chức năng (SRS)",
            "opening": "Yêu cầu {{MaYeuCau}}: Khi {{DieuKienKichHoat}}, hệ thống bắt buộc thực thi {{HanhVi}} và trả về {{KetQuaDauRa}}.",
            "basis": "Giới hạn vận hành: thời gian phản hồi tối đa <= {{LatencyMs}} ms; mức tiêu thụ bộ nhớ <= {{MaxRamMb}} MB.",
            "transition": "Xử lý ngoại lệ: nếu đầu vào không hợp lệ hoặc dịch vụ phụ thuộc ngoại tuyến, hệ thống phải kích hoạt cơ chế dự phòng {{Fallback}} và ghi log mã lỗi chuẩn.",
            "closing": "Tiêu chí nghiệm thu: 100% ca kiểm thử đơn vị và tích hợp liên quan vượt qua không phát sinh hồi quy.",
        },
    },
}


def get_phrase_templates(
    style_id: str = "hanh_chinh",
    doc_type: Optional[str] = None,
) -> Dict[str, Any]:
    """Return curated Vietnamese phrasing templates for a style and optional document subtype."""
    canonical = resolve_style_id(style_id)
    templates = PHRASE_TEMPLATES.get(canonical, {})
    if doc_type:
        key = str(doc_type).strip().lower()
        if key not in templates:
            raise ValueError(
                f"Không tìm thấy mẫu câu '{doc_type}' trong văn phong '{canonical}'. "
                f"Các loại hợp lệ: {', '.join(templates.keys()) or 'không có'}."
            )
        return {
            "style_id": canonical,
            "doc_type": key,
            "template": templates[key],
        }
    return {
        "style_id": canonical,
        "available_doc_types": list(templates.keys()),
        "templates": templates,
    }


# ---------------------------------------------------------------------------
# Masking & Typography Helpers
# ---------------------------------------------------------------------------

_PROTECTED_SPAN_RE = re.compile(
    r"```[\s\S]*?```"        # fenced code blocks
    r"|`[^`\n]+`"            # inline code
    r"|\{\{[^}\n]+\}\}"      # {{TOKEN}} placeholders
    r"|\{%[^%\n]+%\}"        # {% tag %}
    r"|https?://[^\s\)\]>\"']+"  # URLs
)
_QUOTED_SPAN_RE = re.compile(r'“[^”\n]*”|"[^"\n]*"')


def _mask_protected_spans(
    text: str,
    preserve_quotes: bool = True,
) -> Tuple[str, List[str]]:
    """Replace code, placeholders, URLs, and optionally quotes with collision-free sentinel tokens."""
    masked_values: List[str] = []

    def _store(match: re.Match[str]) -> str:
        idx = len(masked_values)
        masked_values.append(match.group(0))
        return f"\uE000WSMASK{idx}\uE001"

    out = _PROTECTED_SPAN_RE.sub(_store, text)
    if preserve_quotes:
        out = _QUOTED_SPAN_RE.sub(_store, out)
    return out, masked_values


def _unmask_protected_spans(text: str, masked_values: Sequence[str]) -> str:
    """Restore sentinel tokens back to their original literal text."""
    out = text
    for idx, original in enumerate(masked_values):
        out = out.replace(f"\uE000WSMASK{idx}\uE001", original)
    return out


def _sanitize_typography(text: str) -> str:
    """Apply Vietnamese NFC & punctuation spacing cleanup, reusing core.docx when available."""
    if not text:
        return text
    try:
        from core.docx.advanced_engine import VietnameseTypographySanitizer

        return VietnameseTypographySanitizer.sanitize_string(text)
    except Exception:
        s = unicodedata.normalize("NFC", text)
        s = re.sub(r"[ \t]+([,.:;!?\)\]\}»”’])", r"\1", s)
        s = re.sub(r"([\(\[\{«“‘])[ \t]+", r"\1", s)
        s = re.sub(r"[ \t]{2,}", " ", s)
        return s


def _match_case(replacement: str, matched: str) -> str:
    """Preserve initial capitalization from the matched text when substituting."""
    if not replacement or not matched:
        return replacement
    first_alpha = next((ch for ch in matched if ch.isalpha()), "")
    if not first_alpha:
        return replacement
    if matched.isupper() and len(matched) > 2 and any(c.islower() for c in replacement):
        return replacement.upper()
    if first_alpha.isupper() and replacement[0].islower():
        return replacement[0].upper() + replacement[1:]
    return replacement


def _rule_applies_to_style(rule: StyleRule, style_id: str) -> bool:
    return not rule.styles or style_id in rule.styles


def _resolve_replacement(rule: StyleRule, style_id: str) -> str:
    if style_id in rule.replacements:
        return rule.replacements[style_id]
    return rule.replacements.get("*", "")


def _compile_rule_regex(rule: StyleRule) -> re.Pattern[str]:
    flags = 0 if rule.case_sensitive else re.IGNORECASE
    pattern = rf"(?<![\w\-/])(?:{rule.pattern})(?![\w\-/])"
    return re.compile(pattern, flags)


_COMPILED_RULES: Tuple[Tuple[StyleRule, re.Pattern[str]], ...] = tuple(
    (rule, _compile_rule_regex(rule)) for rule in STYLE_RULES
)


# ---------------------------------------------------------------------------
# Analysis & Refinement Core
# ---------------------------------------------------------------------------

def _split_sentences(text: str) -> List[str]:
    """Split Vietnamese/English prose into non-empty sentences."""
    clean = re.sub(r"\s+", " ", text).strip()
    if not clean:
        return []
    parts = re.split(r"(?<=[.!?])\s+(?=[A-ZÀ-Ỹ0-9\"“(\[])", clean)
    return [p.strip() for p in parts if p.strip()]


def _count_words(text: str) -> int:
    return len(re.findall(r"[\wÀ-ỹ]+", text))


def analyze_writing_style(
    text: str,
    style_id: str = "hanh_chinh",
    max_sentence_words: Optional[int] = None,
) -> StyleAnalysisResult:
    """Analyze Vietnamese/enterprise prose against a target writing style register."""
    canonical = resolve_style_id(style_id)
    profile = WRITING_STYLE_REGISTRY[canonical]
    normalized = unicodedata.normalize("NFC", text or "")
    sentence_limit = max_sentence_words or profile.max_sentence_words

    paragraphs = [p.strip() for p in re.split(r"\r?\n+", normalized) if p.strip()]
    all_sentences: List[str] = []
    findings: List[StyleFinding] = []

    for p_idx, para in enumerate(paragraphs):
        masked_para, _ = _mask_protected_spans(para, preserve_quotes=True)
        sents = _split_sentences(para)
        all_sentences.extend(sents)

        # Check overly long sentences
        for sent in sents:
            wc = _count_words(sent)
            if wc > sentence_limit:
                snippet = sent[:90] + ("..." if len(sent) > 90 else "")
                findings.append(
                    StyleFinding(
                        rule_id="STYLE_LONG_SENTENCE",
                        category="readability",
                        severity="warning",
                        matched_text=snippet,
                        suggestion=f"Tách câu thành 2–3 câu ngắn hơn (<= {sentence_limit} từ/câu).",
                        message=f"Câu dài {wc} từ, vượt ngưỡng khuyến nghị ({sentence_limit} từ) cho văn phong {profile.name_vi}.",
                        paragraph_index=p_idx,
                    )
                )

        # Check duplicate adjacent words
        for m in _DUP_WORD_REGEX.finditer(masked_para):
            findings.append(
                StyleFinding(
                    rule_id="STYLE_DUPLICATE_WORD",
                    category="redundancy",
                    severity="warning",
                    matched_text=m.group(0),
                    suggestion=m.group(1),
                    message=f"Lặp từ liên tiếp '{m.group(0)}'.",
                    paragraph_index=p_idx,
                    start=m.start(),
                    end=m.end(),
                )
            )

        # Table-driven rules
        for rule, regex in _COMPILED_RULES:
            if not _rule_applies_to_style(rule, canonical):
                continue
            repl_preview = _resolve_replacement(rule, canonical)
            for m in regex.finditer(masked_para):
                findings.append(
                    StyleFinding(
                        rule_id=rule.rule_id,
                        category=rule.category,
                        severity=rule.severity,
                        matched_text=m.group(0),
                        suggestion=_match_case(repl_preview, m.group(0)) if repl_preview else "(lược bỏ)",
                        message=rule.message,
                        paragraph_index=p_idx,
                        start=m.start(),
                        end=m.end(),
                    )
                )

    total_words = _count_words(normalized)
    total_sentences = max(len(all_sentences), 1 if total_words > 0 else 0)
    avg_sentence_words = round(total_words / total_sentences, 2) if total_sentences else 0.0
    long_sentences_count = sum(1 for f in findings if f.rule_id == "STYLE_LONG_SENTENCE")

    # Connector presence ratio
    lowered = normalized.lower()
    connectors_found = [c for c in profile.recommended_connectors if c.lower() in lowered]

    # Compute suitability score (100 - weighted deductions)
    severity_weights = {"error": 10, "warning": 5, "info": 2}
    deduction = sum(severity_weights.get(f.severity, 3) for f in findings)
    score = max(0, min(100, 100 - deduction)) if total_words > 0 else 100

    recommendations: List[str] = []
    if any(f.category in ("shorthand", "colloquial") for f in findings):
        recommendations.append(
            f"Chuẩn hóa các từ viết tắt/khẩu ngữ về ngôn ngữ chuẩn cho văn phong {profile.name_vi}."
        )
    if any(f.category == "pronoun" for f in findings):
        recommendations.append(
            f"Thay thế đại từ thân mật bằng chủ thể chuẩn: {', '.join(profile.preferred_pronouns[:3])}."
        )
    if long_sentences_count > 0:
        recommendations.append(
            f"Có {long_sentences_count} câu vượt quá {sentence_limit} từ; nên tách câu để tăng độ mạch lạc."
        )
    if total_sentences >= 3 and not connectors_found:
        recommendations.append(
            f"Cân nhắc bổ sung từ nối chuyển ý phù hợp ({', '.join(profile.recommended_connectors[:4])})."
        )
    if not recommendations:
        recommendations.append(f"Văn phong hiện tại đạt chuẩn tốt cho ngữ vực {profile.name_vi}.")

    return StyleAnalysisResult(
        style_id=canonical,
        style_name=profile.name_vi,
        suitability_score=score,
        metrics={
            "paragraph_count": len(paragraphs),
            "sentence_count": total_sentences,
            "word_count": total_words,
            "avg_sentence_words": avg_sentence_words,
            "long_sentences_count": long_sentences_count,
            "connectors_detected": connectors_found,
        },
        findings=findings,
        recommendations=recommendations,
    )


def _refine_unmasked_segment(
    segment: str,
    style_id: str,
    mode: str,
) -> Tuple[str, List[Dict[str, Any]]]:
    """Apply deterministic table-driven edits and duplicate-word cleanup to an unmasked segment."""
    if not segment:
        return segment, []

    edits: List[Dict[str, Any]] = []
    working = segment

    # Collapse duplicated adjacent words first
    def _dup_repl(m: re.Match[str]) -> str:
        edits.append({"rule_id": "STYLE_DUPLICATE_WORD", "from": m.group(0), "to": m.group(1)})
        return m.group(1)

    working = _DUP_WORD_REGEX.sub(_dup_repl, working)

    for rule, regex in _COMPILED_RULES:
        if rule.analyze_only:
            continue
        if not _rule_applies_to_style(rule, style_id):
            continue
        if mode == "safe" and rule.min_mode == "strict":
            continue

        target_repl = _resolve_replacement(rule, style_id)

        def _apply(m: re.Match[str], _rule: StyleRule = rule, _tpl: str = target_repl) -> str:
            rep = _match_case(_tpl, m.group(0)) if _tpl else ""
            edits.append({"rule_id": _rule.rule_id, "from": m.group(0), "to": rep})
            return rep

        working = regex.sub(_apply, working)

    # Clean orphan whitespace left after stripping trailing particles (e.g. "xem giúp ,")
    working = re.sub(r"[ \t]+([,.;!?])", r"\1", working)
    working = re.sub(r"[ \t]{2,}", " ", working)
    # Preserve leading/trailing newline structure while tidying line-level spaces
    lines = [line.rstrip(" \t") for line in working.splitlines()]
    had_trailing_nl = segment.endswith("\n")
    working = "\n".join(lines) + ("\n" if had_trailing_nl and lines else "")
    working = _sanitize_typography(working)
    return working, edits


def refine_writing_style(
    text: str,
    style_id: str = "hanh_chinh",
    mode: str = "safe",
    preserve_quotes: bool = True,
    include_analysis: bool = True,
) -> StyleRefinementResult:
    """
    Deterministically refine prose style offline while protecting placeholders, code, URLs, and quotes.

    Args:
        text: Source plain text or Markdown prose.
        style_id: Target writing style (`hanh_chinh`, `hoc_thuat`, `doanh_nghiep`, `phap_ly`, `ky_thuat`).
        mode: `'safe'` (shorthand, colloquial, wordiness, typos) or `'strict'` (+ informal pronouns).
        preserve_quotes: Keep quoted speech/citations unmodified when True.
        include_analysis: Compute before/after `StyleAnalysisResult` metrics.
    """
    canonical = resolve_style_id(style_id)
    profile = WRITING_STYLE_REGISTRY[canonical]
    clean_mode = str(mode or "safe").strip().lower()
    if clean_mode not in ("safe", "strict"):
        raise ValueError("Tham số 'mode' chỉ nhận 'safe' hoặc 'strict'.")

    raw_text = text or ""
    before_dict = analyze_writing_style(raw_text, style_id=canonical).to_dict() if include_analysis else None

    masked, spans = _mask_protected_spans(raw_text, preserve_quotes=preserve_quotes)
    transformed_masked, edits = _refine_unmasked_segment(masked, style_id=canonical, mode=clean_mode)
    refined = _unmask_protected_spans(transformed_masked, spans)

    after_dict = analyze_writing_style(refined, style_id=canonical).to_dict() if include_analysis else None

    return StyleRefinementResult(
        success=True,
        style_id=canonical,
        style_name=profile.name_vi,
        mode=clean_mode,
        original_text=raw_text,
        refined_text=refined,
        changed=(refined != raw_text),
        changes_count=len(edits) + (1 if refined != raw_text and not edits else 0),
        applied_edits=edits,
        analysis_before=before_dict,
        analysis_after=after_dict,
        ai_used=False,
        fallback_used=False,
        provider="deterministic_offline",
    )


# ---------------------------------------------------------------------------
# Hybrid AI Style Rewriting (Copilot Integration with Offline Fallback)
# ---------------------------------------------------------------------------

def build_style_copilot_prompt(
    style_id: str = "hanh_chinh",
    document_type: Optional[str] = None,
    instructions: Optional[str] = None,
) -> str:
    """Construct a strict system prompt instructing an LLM to rewrite text in the target register."""
    profile = get_writing_style_profile(style_id)
    lines = [
        f"Bạn là chuyên gia biên tập hành văn tiếng Việt cho ngữ vực **{profile.name_vi}**.",
        profile.ai_system_directive,
        f"Giọng điệu chủ đạo: {', '.join(profile.tone_keywords)}.",
        f"Xưng hô chuẩn: {', '.join(profile.preferred_pronouns)}.",
        f"Ràng buộc: {profile.forbidden_patterns_note}",
        "BẮT BUỘC: Bảo toàn 100% các placeholder `{{...}}`, đoạn mã `` `...` ``, số liệu, ngày tháng, tên riêng và trích dẫn nguyên văn; chỉ trả về văn bản đã biên tập hoàn chỉnh, không thêm lời giải thích dư thừa.",
    ]
    if document_type:
        lines.append(f"Loại văn bản mục tiêu: {document_type}.")
    if instructions and instructions.strip():
        lines.append(f"Yêu cầu bổ sung từ người dùng: {instructions.strip()}")
    return "\n".join(lines)


def rewrite_with_style(
    text: str,
    style_id: str = "hanh_chinh",
    mode: str = "safe",
    use_ai: bool = False,
    copilot_mode: str = "local",
    document_type: Optional[str] = None,
    instructions: Optional[str] = None,
    preserve_quotes: bool = True,
    context_chunks: Optional[List[Dict[str, Any]]] = None,
) -> StyleRefinementResult:
    """
    Rewrite prose using Hybrid AI (when `use_ai=True` and endpoint is online),
    with deterministic offline refinement as guaranteed zero-failure fallback.
    """
    canonical = resolve_style_id(style_id)
    profile = WRITING_STYLE_REGISTRY[canonical]
    raw_text = text or ""

    if not use_ai or not raw_text.strip():
        return refine_writing_style(
            raw_text,
            style_id=canonical,
            mode=mode,
            preserve_quotes=preserve_quotes,
        )

    before_dict = analyze_writing_style(raw_text, style_id=canonical).to_dict()
    sys_prompt = build_style_copilot_prompt(
        style_id=canonical,
        document_type=document_type,
        instructions=instructions,
    )

    try:
        from core.copilot.copilot_provider import CopilotOrchestrator

        orch = CopilotOrchestrator()
        resp = orch.chat(
            prompt=f"Hãy viết lại đoạn văn bản sau theo đúng văn phong yêu cầu:\n\n{raw_text}",
            context_chunks=context_chunks,
            mode=copilot_mode,
            system_prompt=sys_prompt,
        )
        if resp.success and not resp.fallback_used and (resp.answer or "").strip():
            # Post-check AI output with safe deterministic typography & rule cleanup
            post = refine_writing_style(
                resp.answer.strip(),
                style_id=canonical,
                mode="safe",
                preserve_quotes=preserve_quotes,
                include_analysis=False,
            )
            after_dict = analyze_writing_style(post.refined_text, style_id=canonical).to_dict()
            return StyleRefinementResult(
                success=True,
                style_id=canonical,
                style_name=profile.name_vi,
                mode=mode,
                original_text=raw_text,
                refined_text=post.refined_text,
                changed=(post.refined_text != raw_text),
                changes_count=max(1, post.changes_count) if post.refined_text != raw_text else 0,
                applied_edits=post.applied_edits,
                analysis_before=before_dict,
                analysis_after=after_dict,
                ai_used=True,
                fallback_used=False,
                provider=f"{resp.provider}:{resp.model}",
            )
    except Exception as exc:
        ai_err = str(exc)
    else:
        ai_err = "LLM offline or unavailable; fell back to deterministic writing-style refinement."

    fallback_res = refine_writing_style(
        raw_text,
        style_id=canonical,
        mode=mode,
        preserve_quotes=preserve_quotes,
    )
    fallback_res.fallback_used = True
    fallback_res.provider = "deterministic_offline_fallback"
    fallback_res.error = ai_err
    return fallback_res


# ---------------------------------------------------------------------------
# DOCX OpenXML Run-Safe Integration (ERR_DOCX_002)
# ---------------------------------------------------------------------------

def _paragraph_has_protected_xml(paragraph: Any) -> bool:
    """Return True if paragraph embeds drawings, fields, hyperlinks, or OMML equations."""
    return bool(
        paragraph._p.xpath(
            ".//*[local-name()='drawing' or local-name()='pict' or local-name()='fldChar' "
            "or local-name()='fldSimple' or local-name()='hyperlink' or local-name()='oMath']"
        )
    )


def _runs_have_uniform_style(runs: Sequence[Any]) -> bool:
    """Check whether all non-empty runs share identical basic formatting."""
    signatures = set()
    for r in runs:
        if not r.text:
            continue
        color_rgb = None
        try:
            color_rgb = str(r.font.color.rgb) if r.font.color and r.font.color.rgb else None
        except Exception:
            color_rgb = None
        signatures.add((bool(r.bold), bool(r.italic), bool(r.underline), r.font.name, r.font.size, color_rgb))
        if len(signatures) > 1:
            return False
    return True


def _apply_transform_to_paragraph(
    paragraph: Any,
    transform_fn: Callable[[str], StyleRefinementResult],
) -> Tuple[int, List[Dict[str, Any]]]:
    """
    Apply prose transformation to a Word paragraph while preserving run-level formatting (ERR_DOCX_002).
    Returns `(modified_runs_or_paragraphs_count, applied_edits)`.
    """
    runs = list(getattr(paragraph, "runs", []))
    if not runs:
        return 0, []

    edits: List[Dict[str, Any]] = []
    has_protected = _paragraph_has_protected_xml(paragraph)

    # Single-run or uniform-style plain paragraph: transform full paragraph text safely
    if not has_protected and (len(runs) == 1 or _runs_have_uniform_style(runs)):
        full_text = "".join(r.text for r in runs)
        if not full_text.strip():
            return 0, []
        res = transform_fn(full_text)
        if res.refined_text != full_text:
            runs[0].text = res.refined_text
            for extra_r in runs[1:]:
                extra_r.text = ""
            return 1, res.applied_edits
        return 0, []

    # Mixed formatting or protected XML nodes: transform strictly run-by-run
    modified = 0
    for r in runs:
        if not r.text or not r.text.strip():
            continue
        res = transform_fn(r.text)
        if res.refined_text != r.text:
            r.text = res.refined_text
            modified += 1
            edits.extend(res.applied_edits)
    return modified, edits


def _is_heading_paragraph(paragraph: Any) -> bool:
    style_name = (paragraph.style.name or "").strip().lower() if getattr(paragraph, "style", None) else ""
    return style_name.startswith("heading") or style_name in ("title", "subtitle")


def analyze_docx_writing_style(
    input_path: str,
    style_id: str = "hanh_chinh",
    include_tables: bool = True,
) -> Dict[str, Any]:
    """Extract body (and optional table) text from a `.docx` file and return writing-style analysis."""
    src = Path(input_path)
    if not src.is_file():
        raise FileNotFoundError(f"DOCX file not found: {input_path}")

    from docx import Document

    doc = Document(str(src))
    segments: List[str] = [p.text for p in doc.paragraphs if (p.text or "").strip()]
    if include_tables:
        for tbl in doc.tables:
            for row in tbl.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        if (p.text or "").strip():
                            segments.append(p.text)

    combined = "\n".join(segments)
    report = analyze_writing_style(combined, style_id=style_id).to_dict()
    report["input_path"] = str(src)
    report["paragraphs_scanned"] = len(segments)
    return report


def apply_writing_style_to_document(
    doc: Any,
    style_id: str = "hanh_chinh",
    mode: str = "safe",
    preserve_quotes: bool = True,
    include_headings: bool = False,
    include_tables: bool = True,
) -> Dict[str, Any]:
    """
    Apply deterministic prose style normalization to an in-memory `docx.Document`
    preserving all run-level formatting, tables, images, and OpenXML invariants.
    """
    canonical = resolve_style_id(style_id)

    def _xform(txt: str) -> StyleRefinementResult:
        return refine_writing_style(
            txt,
            style_id=canonical,
            mode=mode,
            preserve_quotes=preserve_quotes,
            include_analysis=False,
        )

    paragraphs_modified = 0
    all_edits: List[Dict[str, Any]] = []
    scanned_texts: List[str] = []

    for p in doc.paragraphs:
        if include_headings or not _is_heading_paragraph(p):
            mod, ed = _apply_transform_to_paragraph(p, _xform)
            paragraphs_modified += mod
            all_edits.extend(ed)
        if (p.text or "").strip():
            scanned_texts.append(p.text)

    if include_tables:
        for tbl in doc.tables:
            for row in tbl.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        if include_headings or not _is_heading_paragraph(p):
                            mod, ed = _apply_transform_to_paragraph(p, _xform)
                            paragraphs_modified += mod
                            all_edits.extend(ed)
                        if (p.text or "").strip():
                            scanned_texts.append(p.text)

    post_report = analyze_writing_style("\n".join(scanned_texts), style_id=canonical).to_dict()
    return {
        "style_id": canonical,
        "style_name": WRITING_STYLE_REGISTRY[canonical].name_vi,
        "mode": mode,
        "paragraphs_modified": paragraphs_modified,
        "edits_count": len(all_edits),
        "applied_edits": all_edits[:100],
        "suitability_score": post_report["suitability_score"],
        "total_findings_remaining": post_report["total_findings"],
    }


def apply_writing_style_to_docx(
    input_path: str,
    output_path: Optional[str] = None,
    style_id: str = "hanh_chinh",
    mode: str = "safe",
    preserve_quotes: bool = True,
    include_headings: bool = False,
    include_tables: bool = True,
) -> Dict[str, Any]:
    """
    Apply deterministic prose style normalization to an existing Word `.docx` document
    preserving all run-level formatting, tables, images, and OpenXML invariants.
    """
    src = Path(input_path)
    if not src.is_file():
        raise FileNotFoundError(f"DOCX file not found: {input_path}")

    canonical = resolve_style_id(style_id)
    dst = Path(output_path) if output_path else src.with_name(f"{src.stem}_{canonical}.docx")
    dst.parent.mkdir(parents=True, exist_ok=True)

    from docx import Document
    from core.docx.advanced_engine import safe_save_docx

    doc = Document(str(src))
    style_stats = apply_writing_style_to_document(
        doc,
        style_id=canonical,
        mode=mode,
        preserve_quotes=preserve_quotes,
        include_headings=include_headings,
        include_tables=include_tables,
    )
    saved_path = safe_save_docx(doc, str(dst))

    return {
        "ok": True,
        "input_path": str(src),
        "output_path": saved_path,
        **style_stats,
    }


__all__ = [
    "StyleProfile",
    "StyleRule",
    "StyleFinding",
    "StyleAnalysisResult",
    "StyleRefinementResult",
    "WRITING_STYLE_REGISTRY",
    "STYLE_RULES",
    "PHRASE_TEMPLATES",
    "resolve_style_id",
    "get_writing_style_profile",
    "list_writing_styles",
    "get_phrase_templates",
    "analyze_writing_style",
    "refine_writing_style",
    "build_style_copilot_prompt",
    "rewrite_with_style",
    "analyze_docx_writing_style",
    "apply_writing_style_to_document",
    "apply_writing_style_to_docx",
]
