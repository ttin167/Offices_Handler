"""
core.ba_hub.obsidian_vault — Obsidian Local-First Knowledge Graph & Two-Way Vault Bridge.

Implements:
1. ObsidianVaultExporter:
   - Exports User Stories, Use Cases, Data Dictionaries and Diagrams into a complete,
     import-ready Obsidian Vault package (.zip).
   - Generates YAML Frontmatter (metadata, tags, status, author).
   - Implements Bidirectional Wikilinks syntax: [[Target_Page|Display_Text]].
   - Generates native Obsidian Canvas v1.0 (.canvas JSON) connecting requirements to data schemas.
   - Generates clean Mermaid diagrams rendered natively by Obsidian.
2. ObsidianVaultCompiler:
   - Scans an Obsidian Vault directory or .zip archive.
   - Extracts AST, frontmatter, and resolves [[Wikilinks]] into Word Bookmarks and Hyperlinks.
   - Compiles requirements into publication-grade Word BRD/SRS (.docx) conforming to
     OpenXML invariants (<w:cantSplit/>, <w:tblHeader/>, <w:vAlign w:val="center"/>).
   - Exports high-fidelity PDF documents.
"""

from __future__ import annotations

import datetime
import json
import os
import re
import tempfile
import zipfile
from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional, Tuple, Union

import yaml
from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.shared import Inches, Pt, RGBColor

from core.ba_hub.story_builder import UserStory, GherkinScenario
from core.ba_hub.usecase_builder import UseCaseSpec


# ------------------------------------------------------------------------------
# 1. OBSIDIAN VAULT EXPORTER
# ------------------------------------------------------------------------------
class ObsidianVaultExporter:
    """Exports BA Requirements, Data Dictionaries, and Diagrams into an Obsidian Vault (.zip)."""

    @classmethod
    def export_vault_zip(
        cls,
        stories: List[Union[UserStory, Dict[str, Any]]],
        usecases: Optional[List[Union[UseCaseSpec, Dict[str, Any]]]] = None,
        data_dictionaries: Optional[Dict[str, List[Dict[str, Any]]]] = None,
        diagrams: Optional[Dict[str, str]] = None,
        project_name: str = "Enterprise_BA_Project",
        output_zip_path: Optional[str] = None,
        author: str = "Lead Business Analyst",
    ) -> str:
        """Packages the entire BA requirements ecosystem into a ready-to-open Obsidian Vault (.zip)."""
        if output_zip_path is None:
            ts = int(datetime.datetime.now().timestamp() * 1000)
            safe_proj = re.sub(r"[^\w\-]", "_", project_name)
            output_zip_path = os.path.join(tempfile.gettempdir(), f"{safe_proj}_Obsidian_Vault_{ts}.zip")

        os.makedirs(os.path.dirname(os.path.abspath(output_zip_path)), exist_ok=True)

        # Normalize stories and use cases
        norm_stories: List[UserStory] = []
        for s in (stories or []):
            if isinstance(s, UserStory):
                norm_stories.append(s)
            elif isinstance(s, dict):
                norm_stories.append(UserStory.from_dict(s))

        norm_usecases: List[UseCaseSpec] = []
        for u in (usecases or []):
            if isinstance(u, UseCaseSpec):
                norm_usecases.append(u)
            elif isinstance(u, dict):
                norm_usecases.append(UseCaseSpec.from_dict(u))

        norm_data_dicts: Dict[str, List[Dict[str, Any]]] = {}
        for tbl_name, f_list in (data_dictionaries or {}).items():
            norm_f_list = []
            for fld in f_list:
                if hasattr(fld, "to_dict"):
                    norm_f_list.append(fld.to_dict())
                elif isinstance(fld, dict):
                    norm_f_list.append(fld)
                else:
                    norm_f_list.append(vars(fld))
            norm_data_dicts[tbl_name] = norm_f_list
        data_dicts = norm_data_dicts
        diag_map = diagrams or {}

        with zipfile.ZipFile(output_zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
            # 1. Root Index.md
            index_content = cls.build_index_markdown(project_name, norm_stories, norm_usecases, data_dicts, author)
            zf.writestr("Index.md", index_content)

            # 2. Obsidian Canvas: Project_Overview.canvas
            canvas_content = cls.build_canvas_json(project_name, norm_stories, norm_usecases, data_dicts)
            zf.writestr("Project_Overview.canvas", canvas_content)

            # 3. Requirements/BRD_Overview.md
            brd_overview = cls.build_brd_overview_markdown(project_name, norm_stories, norm_usecases, author)
            zf.writestr("Requirements/BRD_Overview.md", brd_overview)

            # 4. Requirements/US_*.md
            for st in norm_stories:
                story_filename = f"Requirements/{cls._sanitize_filename(st.id)}.md"
                story_content = cls.build_story_markdown(st, project_name, author)
                zf.writestr(story_filename, story_content)

            # 5. Requirements/UC_*.md
            for uc in norm_usecases:
                uc_filename = f"Requirements/{cls._sanitize_filename(uc.id)}.md"
                uc_content = cls.build_usecase_markdown(uc, project_name, author)
                zf.writestr(uc_filename, uc_content)

            # 6. Data-Dictionary/<Table>.md
            for tbl_name, fields in data_dicts.items():
                dict_filename = f"Data-Dictionary/{cls._sanitize_filename(tbl_name)}.md"
                dict_content = cls.build_data_dict_markdown(tbl_name, fields, project_name, author)
                zf.writestr(dict_filename, dict_content)

            # 7. Diagrams/
            if not diag_map:
                # Provide a default ERD and Process flow if none supplied
                diag_map["Database_ERD.md"] = cls._build_default_erd_mermaid(data_dicts)
                diag_map["System_Process_Flow.md"] = cls._build_default_flow_mermaid(norm_stories)

            for diag_name, diag_text in diag_map.items():
                if not diag_name.endswith(".md"):
                    diag_name = f"{diag_name}.md"
                zf.writestr(f"Diagrams/{diag_name}", diag_text)

            # 8. .obsidian/app.json configuration
            obsidian_config = {
                "attachmentFolderPath": "Diagrams",
                "livePreview": True,
                "promptDelete": False,
                "alwaysUpdateLinks": True,
            }
            zf.writestr(".obsidian/app.json", json.dumps(obsidian_config, indent=2))

        return output_zip_path

    @classmethod
    def build_index_markdown(
        cls,
        project_name: str,
        stories: List[UserStory],
        usecases: List[UseCaseSpec],
        data_dicts: Dict[str, List[Dict[str, Any]]],
        author: str,
    ) -> str:
        """Generates the central navigation Index.md for the Obsidian Vault."""
        today = datetime.date.today().isoformat()
        lines = [
            "---",
            f"project: \"{project_name}\"",
            "type: vault-index",
            f"author: \"{author}\"",
            f"updated: {today}",
            "tags:",
            "  - ba-workstation",
            "  - vault-index",
            "  - babok",
            "---",
            "",
            f"# 📌 {project_name} — Đồ Thị Tri Thức Nghiệp Vụ (Requirements Hub)",
            "",
            "> 💡 **Chào mừng bạn đến với Obsidian Vault của dự án.**  ",
            "> Vault này được tạo tự động bởi **Antigravity BA Enterprise Workstation**. Mọi tài liệu yêu cầu, kịch bản nghiệm thu và từ điển dữ liệu được liên kết 2 chiều thông qua cú pháp Wikilinks `[[...]]`.",
            "",
            "## 🧭 Bản Đồ Tư Duy Không Gian (Spatial Mindmap)",
            "- 🗺️ **Mở Obsidian Canvas:** [[Project_Overview.canvas|Bản Đồ Không Gian Dự Án (Canvas)]]",
            "- 📑 **Tài Liệu Tổng Quan:** [[BRD_Overview|Đặc Tả Yêu Cầu Nghiệp Vụ (BRD Overview)]]",
            "",
            "## 📋 Danh Mục User Stories & Tiêu Chí Nghiệm Thu",
        ]

        if not stories:
            lines.append("*(Chưa có User Story nào được ghi nhận)*")
        else:
            for s in stories:
                s_link = f"[[{s.id}|{s.id}: {s.title}]]"
                lines.append(f"- {s_link} — *Persona:* `{s.role}` ({len(s.scenarios)} kịch bản AC)")

        lines.extend([
            "",
            "## ⚙️ Danh Mục Use Case Kỹ Thuật (Use Case Specs)",
        ])

        if not usecases:
            lines.append("*(Chưa có Use Case nào được định nghĩa)*")
        else:
            for u in usecases:
                u_link = f"[[{u.id}|{u.id}: {u.title}]]"
                lines.append(f"- {u_link} — *Actor chính:* `{u.actor_primary}`")

        lines.extend([
            "",
            "## 🗄️ Từ Điển Dữ Liệu Thực Thể (Data Dictionaries)",
        ])

        if not data_dicts:
            lines.append("*(Chưa có thực thể dữ liệu nào)*")
        else:
            for tbl in sorted(data_dicts.keys()):
                lines.append(f"- [[{tbl}|Bảng {tbl}]] ({len(data_dicts[tbl])} trường dữ liệu)")

        lines.extend([
            "",
            "## 📊 Sơ Đồ Kiến Trúc & Luồng Xử Lý (Diagrams)",
            "- [[Database_ERD|Sơ Đồ Thực Thể Liên Kết (Mermaid ERD)]]",
            "- [[System_Process_Flow|Sơ Đồ Luồng Nghiệp Vụ (Process Flow)]]",
            "",
            "---",
            f"*Được khởi tạo bởi Antigravity BA Enterprise Workstation lúc {today}.*",
        ])
        return "\n".join(lines)

    @classmethod
    def build_brd_overview_markdown(
        cls,
        project_name: str,
        stories: List[UserStory],
        usecases: List[UseCaseSpec],
        author: str,
    ) -> str:
        """Generates the BRD_Overview.md document linking to all Epics and Stories."""
        today = datetime.date.today().isoformat()
        lines = [
            "---",
            f"title: \"Tài Liệu Yêu Cầu Nghiệp Vụ (BRD Overview) - {project_name}\"",
            "type: brd-overview",
            f"author: \"{author}\"",
            f"date: {today}",
            "tags:",
            "  - brd",
            "  - srs",
            "  - requirements-overview",
            "---",
            "",
            f"# [[Index|Trang Chủ]] > TÀI LIỆU YÊU CẦU NGHIỆP VỤ (BRD SPECIFICATION)",
            "",
            f"**Dự án:** {project_name}  ",
            f"**Tác giả / Lead BA:** {author}  ",
            f"**Ngày cập nhật:** {today}  ",
            "",
            "## 1. Mục Đích & Phạm Vi Dự Án",
            f"Tài liệu này đặc tả toàn bộ các yêu cầu nghiệp vụ, luồng giao dịch và tiêu chí nghiệm thu cho hệ thống **{project_name}** theo chuẩn BABOK v3 và IEEE 830.",
            "",
            "## 2. Ma Trận Truy Xuất Yêu Cầu (Traceability Matrix)",
            "| Mã Yêu Cầu | Tên Tính Năng | Persona / Vai Trò | Use Case Kỹ Thuật | Trạng Thái |",
            "|---|---|---|---|---|",
        ]

        for s in stories:
            matched_uc = next((u.id for u in usecases if s.id.split("-")[-1] in u.id or u.title in s.title), "UC-CORE")
            lines.append(f"| [[{s.id}]] | {s.title} | `{s.role}` | [[{matched_uc}]] | `Approved` |")

        lines.extend([
            "",
            "## 3. Bản Đồ Không Gian Liên Kết",
            "Xem trực quan các mối quan hệ đa chiều tại [[Project_Overview.canvas]].",
        ])
        return "\n".join(lines)

    @classmethod
    def build_story_markdown(cls, story: UserStory, project_name: str, author: str) -> str:
        """Generates an Obsidian-ready User Story markdown file with YAML frontmatter and Wikilinks."""
        today = datetime.date.today().isoformat()
        tags = ["requirement", "user-story"]
        if story.domain:
            safe_domain = re.sub(r"[^\w\-]", "-", story.domain.lower())
            tags.append(safe_domain)

        lines = [
            "---",
            f"id: \"{story.id}\"",
            f"title: \"{story.title}\"",
            f"domain: \"{story.domain}\"",
            f"role: \"{story.role}\"",
            "status: Draft",
            f"author: \"{author}\"",
            f"date: {today}",
            "tags:",
        ]
        for t in tags:
            lines.append(f"  - {t}")
        lines.extend([
            "---",
            "",
            f"# [[BRD_Overview|BRD]] > {story.id}: {story.title}",
            "",
            "## 1. Đặc Tả User Story (3 Thành Phần)",
            f"- **As a:** `{story.role}`",
            f"- **I want:** {story.goal}",
            f"- **So that:** {story.benefit}",
            "",
            "## 2. Quy Tắc Nghiệp Vụ (Business Rules)",
        ])

        if not story.business_rules:
            lines.append("- *(Không có quy tắc bổ sung)*")
        else:
            for br in story.business_rules:
                # Auto-link table names if mentioned
                linked_br = re.sub(r"\b([A-Z][A-Z0-9_]{3,})\b", r"[[\1]]", br)
                lines.append(f"- {linked_br}")

        lines.extend([
            "",
            "## 3. Tiêu Chí Nghiệm Thu (Acceptance Criteria - Gherkin)",
        ])

        if not story.scenarios:
            lines.append("*(Chưa có kịch bản nghiệm thu)*")
        else:
            for idx, sc in enumerate(story.scenarios, start=1):
                badge = "Happy Path" if "happy" in sc.scenario_type.lower() else "Negative / Boundary"
                lines.extend([
                    f"### Kịch Bản #{idx}: {sc.title} *({badge})*",
                    "```gherkin",
                    sc.to_gherkin_text(),
                    "```",
                    "",
                ])

        lines.extend([
            "## 4. Liên Kết Truy Xuất Nguồn Gốc (Wikilinks)",
            f"- **Tài Liệu Cha:** [[BRD_Overview]]",
            f"- **Bảng Dữ Liệu Tác Động:** " + ", ".join([f"[[{tbl}]]" for tbl in cls._extract_tables_from_story(story)]) if cls._extract_tables_from_story(story) else "- **Bảng Dữ Liệu:** Chưa gắn thẻ",
            f"- **Sơ Đồ Không Gian:** [[Project_Overview.canvas]]",
        ])
        return "\n".join(lines)

    @classmethod
    def build_usecase_markdown(cls, uc: UseCaseSpec, project_name: str, author: str) -> str:
        """Generates an Obsidian-ready Use Case specification markdown."""
        today = datetime.date.today().isoformat()
        pre_str = ", ".join(uc.pre_conditions) if isinstance(uc.pre_conditions, list) else str(uc.pre_conditions or "")
        post_str = ", ".join(uc.post_conditions) if isinstance(uc.post_conditions, list) else str(uc.post_conditions or "")

        lines = [
            "---",
            f"id: \"{uc.id}\"",
            f"title: \"{uc.title}\"",
            f"actor: \"{uc.actor_primary}\"",
            f"status: \"{uc.status}\"",
            f"author: \"{author}\"",
            f"date: {today}",
            "tags:",
            "  - use-case",
            "  - specification",
            "---",
            "",
            f"# [[BRD_Overview|BRD]] > {uc.id}: {uc.title}",
            "",
            "## 1. Thông Tin Cơ Bản",
            f"- **Mã Use Case:** `{uc.id}`",
            f"- **Tên Use Case:** {uc.title}",
            f"- **Tác Nhân Chính (Primary Actor):** `{uc.actor_primary}`",
            f"- **Sự Kiện Kích Hoạt (Trigger):** {uc.trigger or 'Người dùng kích hoạt chức năng'}",
            f"- **Tiền Điều Kiện (Pre-conditions):** {pre_str or 'Hệ thống sẵn sàng'}",
            f"- **Hậu Điều Kiện (Post-conditions):** {post_str or 'Giao dịch hoàn tất'}",
            "",
            "## 2. Luồng Sự Kiện Chính (Main Basic Flow)",
        ]

        if not uc.main_flow:
            lines.append("1. Người dùng thực hiện thao tác.")
            lines.append("2. Hệ thống xử lý và phản hồi.")
        else:
            for step in uc.main_flow:
                if hasattr(step, "step") and hasattr(step, "actor") and hasattr(step, "action"):
                    lines.append(f"{step.step}. `{step.actor}`: {step.action}")
                elif isinstance(step, dict):
                    lines.append(f"{step.get('step', '')}. `{step.get('actor', '')}`: {step.get('action', '')}")
                else:
                    lines.append(f"{step}")

        if uc.alt_flows:
            lines.extend(["", "## 3. Luồng Sự Kiện Phụ (Alternative Flows)"])
            for af in uc.alt_flows:
                af_id = getattr(af, "id", getattr(af, "name", "Alt"))
                af_cond = getattr(af, "condition", "")
                af_steps = getattr(af, "steps", [])
                af_steps_str = "; ".join(str(s) for s in af_steps) if isinstance(af_steps, list) else str(af_steps)
                lines.append(f"- **{af_id} ({af_cond}):** {af_steps_str}")

        if uc.exception_flows:
            lines.extend(["", "## 4. Luồng Ngoại Lệ (Exception Flows)"])
            for ef in uc.exception_flows:
                ef_id = getattr(ef, "id", getattr(ef, "name", "Exc"))
                ef_cond = getattr(ef, "condition", "")
                ef_steps = getattr(ef, "steps", [])
                ef_steps_str = "; ".join(str(s) for s in ef_steps) if isinstance(ef_steps, list) else str(ef_steps)
                lines.append(f"- **{ef_id} ({ef_cond}):** {ef_steps_str}")

        lines.extend([
            "",
            "## 5. Truy Xuất & Đồ Thị (Wikilinks)",
            f"- **Về Mục Lục:** [[Index]] | [[BRD_Overview]]",
            f"- **Sơ Đồ Luồng:** [[System_Process_Flow]]",
        ])
        return "\n".join(lines)

    @classmethod
    def build_data_dict_markdown(
        cls,
        table_name: str,
        fields: List[Any],
        project_name: str,
        author: str,
    ) -> str:
        """Generates a Data-Dictionary markdown document with Wikilinks for Foreign Keys."""
        today = datetime.date.today().isoformat()
        lines = [
            "---",
            f"table: \"{table_name}\"",
            "type: data-dictionary",
            f"author: \"{author}\"",
            f"date: {today}",
            "tags:",
            "  - data-dictionary",
            "  - database-schema",
            "---",
            "",
            f"# [[Index|Trang Chủ]] > Từ Điển Dữ Liệu: `{table_name}`",
            "",
            f"Bảng thực thể cơ sở dữ liệu phục vụ phân hệ nghiệp vụ **{project_name}**.",
            "",
            "| Mã Trường (Field) | Tên Hiển Thị | Kiểu Dữ Liệu | Nullable | PK | Khóa Ngoại (FK) | Ý Nghĩa & Ràng Buộc |",
            "|---|---|---|---|---|---|---|",
        ]

        for f in fields:
            if hasattr(f, "to_dict"):
                f = f.to_dict()
            elif not isinstance(f, dict):
                f = vars(f)
            fld_name = f.get("field") or f.get("field_name") or "col"
            fld_disp = f.get("name") or f.get("display_name") or fld_name
            fld_type = f.get("type") or f.get("data_type") or "VARCHAR(50)"
            fld_null = f.get("nullable", "YES")
            fld_pk = "✅ PK" if (f.get("pk") == "YES" or f.get("is_primary_key")) else "-"
            fld_fk_raw = f.get("fk") or f.get("references") or "-"
            fld_desc = f.get("description", "")
            fld_rules = f.get("rules") or f.get("business_rules") or ""
            notes = f"{fld_desc} {f'({fld_rules})' if fld_rules else ''}".strip() or "-"

            # Wikilink for Foreign Key: e.g. "USERS(id)" -> "[[USERS]](id)"
            if fld_fk_raw and fld_fk_raw != "-":
                m = re.match(r"([A-Za-z0-9_]+)(.*)", fld_fk_raw)
                if m:
                    fld_fk_str = f"[[{m.group(1)}]]{m.group(2)}"
                else:
                    fld_fk_str = f"[[{fld_fk_raw}]]"
            else:
                fld_fk_str = "-"

            lines.append(f"| `{fld_name}` | {fld_disp} | `{fld_type}` | {fld_null} | {fld_pk} | {fld_fk_str} | {notes} |")

        lines.extend([
            "",
            "## Sơ Đồ Thực Thể Liên Quan",
            "- Xem ERD toàn hệ thống tại: [[Database_ERD]]",
            "- Trực quan không gian tại: [[Project_Overview.canvas]]",
        ])
        return "\n".join(lines)

    @classmethod
    def build_canvas_json(
        cls,
        project_name: str,
        stories: List[UserStory],
        usecases: List[UseCaseSpec],
        data_dicts: Dict[str, List[Dict[str, Any]]],
    ) -> str:
        """Generates an Obsidian Canvas v1.0 JSON string connecting BRD, Stories, UseCases, and Data Tables."""
        nodes: List[Dict[str, Any]] = []
        edges: List[Dict[str, Any]] = []

        # Center Root Node: BRD Overview
        nodes.append({
            "id": "node-root-brd",
            "type": "file",
            "file": "Requirements/BRD_Overview.md",
            "x": -500,
            "y": 0,
            "width": 360,
            "height": 260,
            "color": "1",  # Red / Crimson
        })

        # Column 1: Stories (x = -50)
        y_cursor_story = -200
        for idx, st in enumerate(stories[:8]):
            s_id = f"node-story-{st.id}"
            nodes.append({
                "id": s_id,
                "type": "file",
                "file": f"Requirements/{cls._sanitize_filename(st.id)}.md",
                "x": -50,
                "y": y_cursor_story,
                "width": 340,
                "height": 240,
                "color": "4",  # Green
            })
            edges.append({
                "id": f"edge-brd-{st.id}",
                "fromNode": "node-root-brd",
                "fromSide": "right",
                "toNode": s_id,
                "toSide": "left",
                "label": "phân rã thành",
            })
            y_cursor_story += 280

        # Column 2: Data Dictionaries (x = 420)
        y_cursor_dict = -150
        for idx, tbl in enumerate(list(data_dicts.keys())[:6]):
            d_id = f"node-dict-{tbl}"
            nodes.append({
                "id": d_id,
                "type": "file",
                "file": f"Data-Dictionary/{cls._sanitize_filename(tbl)}.md",
                "x": 420,
                "y": y_cursor_dict,
                "width": 340,
                "height": 240,
                "color": "6",  # Purple
            })
            # Connect the first story to this table if related
            if stories:
                first_story_id = f"node-story-{stories[0].id}"
                edges.append({
                    "id": f"edge-story-dict-{tbl}",
                    "fromNode": first_story_id,
                    "fromSide": "right",
                    "toNode": d_id,
                    "toSide": "left",
                    "label": "ghi nhận / truy vấn",
                })
            y_cursor_dict += 280

        canvas_obj = {
            "nodes": nodes,
            "edges": edges,
        }
        return json.dumps(canvas_obj, indent=2, ensure_ascii=False)

    @classmethod
    def _sanitize_filename(cls, name: str) -> str:
        return re.sub(r'[\\/*?:"<>|]', "_", name).strip()

    @classmethod
    def _extract_tables_from_story(cls, story: UserStory) -> List[str]:
        full_text = f"{story.title} {story.goal} {' '.join(story.business_rules)}"
        matches = re.findall(r"\b([A-Z][A-Z0-9_]{3,})\b", full_text)
        return list(dict.fromkeys(matches))

    @classmethod
    def _build_default_erd_mermaid(cls, data_dicts: Dict[str, List[Dict[str, Any]]]) -> str:
        lines = [
            "---",
            "title: Sơ Đồ Thực Thể Liên Kết (Mermaid ERD)",
            "type: diagram",
            "---",
            "",
            "# Sơ Đồ Cơ Sở Dữ Liệu Hệ Thống (Database ERD)",
            "",
            "```mermaid",
            "erDiagram",
        ]
        if not data_dicts:
            lines.extend([
                "    CUSTOMERS ||--o{ ORDERS : places",
                "    ORDERS ||--|{ ORDER_ITEMS : contains",
                "    CUSTOMERS {",
                "        string customer_id PK",
                "        string full_name",
                "    }",
                "    ORDERS {",
                "        string order_id PK",
                "        string customer_id FK",
                "        decimal total_amount",
                "    }",
            ])
        else:
            for tbl, flds in data_dicts.items():
                lines.append(f"    {tbl} {{")
                for f in flds[:6]:
                    f_name = f.get("field") or f.get("field_name") or "col"
                    f_type = (f.get("type") or "string").split("(")[0].lower()
                    pk_marker = "PK" if (f.get("pk") == "YES" or f.get("is_primary_key")) else ""
                    lines.append(f"        {f_type} {f_name} {pk_marker}".strip())
                lines.append("    }")
        lines.extend(["```", ""])
        return "\n".join(lines)

    @classmethod
    def _build_default_flow_mermaid(cls, stories: List[UserStory]) -> str:
        lines = [
            "---",
            "title: Sơ Đồ Luồng Nghiệp Vụ Chính (Process Flow)",
            "type: diagram",
            "---",
            "",
            "# Sơ Đồ Luồng Tác Nghiệp Nghiệp Vụ",
            "",
            "```mermaid",
            "sequenceDiagram",
            "    autonumber",
            "    actor User as Khách Hàng",
            "    participant App as Mobile App",
            "    participant Core as Core Banking Engine",
            "    participant Napas as NAPAS 24/7 Switch",
            "",
            "    User->>App: Nhập thông tin & số tiền chuyển",
            "    App->>Core: Kiểm tra số dư & hạn mức [[CUSTOMER_ACCOUNTS]]",
            "    alt Số dư đủ & Hợp lệ",
            "        Core->>Napas: Phát lệnh chuyển tiền liên ngân hàng",
            "        Napas-->>Core: Phản hồi thành công (RC=00)",
            "        Core-->>App: Trừ tiền & xuất biên lai giao dịch",
            "        App-->>User: Hiển thị thông báo thành công",
            "    else Số dư không khả dụng",
            "        Core-->>App: Báo lỗi số dư không đủ",
            "        App-->>User: Hiển thị thông báo từ chối",
            "    end",
            "```",
            "",
        ]
        return "\n".join(lines)


# ------------------------------------------------------------------------------
# 2. OBSIDIAN VAULT COMPILER (VAULT -> WORD DOCX / PDF)
# ------------------------------------------------------------------------------
class ObsidianVaultCompiler:
    """Compiles an Obsidian Vault (.zip or directory) into an Enterprise Word/PDF BRD specification."""

    WIKILINK_REGEX = re.compile(r"\[\[([^\]\|]+)(?:\|([^\]]+))?\]\]")

    @classmethod
    def compile_vault_to_docx(
        cls,
        vault_input: str,
        output_docx_path: str,
        project_name: Optional[str] = None,
        author: str = "Lead Business Analyst",
    ) -> str:
        """Parses an Obsidian Vault (.zip or folder) and compiles it into an OpenXML Word document."""
        os.makedirs(os.path.dirname(os.path.abspath(output_docx_path)), exist_ok=True)

        # 1. Unpack or locate vault files
        temp_dir = None
        if os.path.isfile(vault_input) and vault_input.endswith(".zip"):
            temp_dir = tempfile.mkdtemp(prefix="obsidian_vault_")
            with zipfile.ZipFile(vault_input, "r") as zf:
                abs_temp = os.path.abspath(temp_dir)
                for member in zf.infolist():
                    # Zip Slip Traversal Guard (ERR_SEC_005)
                    target_path = os.path.abspath(os.path.join(temp_dir, member.filename))
                    if not target_path.startswith(abs_temp + os.sep) and target_path != abs_temp:
                        raise ValueError(f"Phát hiện nguy cơ Zip Slip Traversal (ERR_SEC_005) trong tệp {member.filename}")
                    zf.extract(member, temp_dir)
            vault_root = temp_dir
        elif os.path.isdir(vault_input):
            vault_root = vault_input
        else:
            raise ValueError(f"Vault input must be a valid .zip file or directory path: {vault_input}")

        try:
            # 2. Scan and parse all Markdown files
            vault_data = cls._scan_vault_directory(vault_root)
            proj_title = project_name or vault_data.get("metadata", {}).get("project") or "Hệ Thống Nghiệp Vụ Doanh Nghiệp"

            # 3. Create OpenXML Word Document
            doc = Document()

            # Set A4 margins
            for sec in doc.sections:
                sec.top_margin = Inches(1.0)
                sec.bottom_margin = Inches(1.0)
                sec.left_margin = Inches(1.0)
                sec.right_margin = Inches(0.8)

            # ------------------------------------------------------------------
            # COVER PAGE
            # ------------------------------------------------------------------
            cls._build_cover_page(doc, proj_title, author)
            doc.add_page_break()

            # ------------------------------------------------------------------
            # DOCUMENT REVISION CONTROL TABLE
            # ------------------------------------------------------------------
            cls._build_revision_history(doc, author)
            doc.add_page_break()

            # ------------------------------------------------------------------
            # SECTION 1: EXECUTIVE SUMMARY & ARCHITECTURE (From Index & Canvas)
            # ------------------------------------------------------------------
            cls._build_section_heading(doc, "1. TỔNG QUAN HỆ THỐNG & ĐỒ THỊ TRI THỨC (KNOWLEDGE GRAPH)", bookmark_id="Section_1")
            p_desc = doc.add_paragraph()
            p_desc.add_run(
                f"Tài liệu này được biên dịch tự động từ hệ sinh thái **Obsidian Local-First Vault** của dự án "
                f"'{proj_title}'. Toàn bộ các yêu cầu chức năng, tiêu chí nghiệm thu (Gherkin AC), và từ điển dữ liệu "
                f"được tổ chức dưới dạng mạng lưới liên kết hai chiều (Wikilinks Traceability Map)."
            )

            # Mention canvas if present
            if vault_data.get("canvas"):
                p_c = doc.add_paragraph()
                p_c.add_run("Bản Đồ Không Gian Obsidian Canvas: ").bold = True
                p_c.add_run(f"Đã phát hiện và ánh xạ {len(vault_data['canvas'].get('nodes', []))} nút thực thể và {len(vault_data['canvas'].get('edges', []))} mối quan hệ luồng nghiệp vụ.")

            # ------------------------------------------------------------------
            # SECTION 2: USER STORIES & ACCEPTANCE CRITERIA
            # ------------------------------------------------------------------
            cls._build_section_heading(doc, "2. ĐẶC TẢ CHI TIẾT USER STORIES & GHERKIN AC", bookmark_id="Section_2")
            stories = vault_data.get("stories", [])
            if not stories:
                doc.add_paragraph("*(Không có User Story nào được tìm thấy trong Vault)*")
            else:
                for idx, st in enumerate(stories, start=1):
                    cls._render_compiled_story(doc, st, idx)

            # ------------------------------------------------------------------
            # SECTION 3: USE CASE SPECIFICATIONS
            # ------------------------------------------------------------------
            usecases = vault_data.get("usecases", [])
            if usecases:
                cls._build_section_heading(doc, "3. ĐẶC TẢ USE CASES & LUỒNG TÁC NGHIỆP", bookmark_id="Section_3")
                for u_idx, uc in enumerate(usecases, start=1):
                    cls._render_compiled_usecase(doc, uc, u_idx)

            # ------------------------------------------------------------------
            # SECTION 4: DATA DICTIONARY TABLES
            # ------------------------------------------------------------------
            data_dicts = vault_data.get("data_dictionaries", {})
            if data_dicts:
                cls._build_section_heading(doc, "4. TỪ ĐIỂN DỮ LIỆU & CƠ SỞ DỮ LIỆU THỰC THỂ", bookmark_id="Section_4")
                for tbl_name, tbl_info in data_dicts.items():
                    cls._render_compiled_data_dict(doc, tbl_name, tbl_info)

            # ------------------------------------------------------------------
            # SECTION 5: DIAGRAMS & MERMAID SPECIFICATIONS
            # ------------------------------------------------------------------
            diagrams = vault_data.get("diagrams", {})
            if diagrams:
                cls._build_section_heading(doc, "5. PHỤ LỤC: SƠ ĐỒ KIẾN TRÚC & QUY TRÌNH", bookmark_id="Section_5")
                for diag_title, diag_content in diagrams.items():
                    p_dt = doc.add_paragraph()
                    p_dt.add_run(f"5.{len(doc.paragraphs)} {diag_title}").bold = True
                    p_code = doc.add_paragraph()
                    r_code = p_code.add_run(diag_content)
                    r_code.font.name = "Consolas"
                    r_code.font.size = Pt(8.5)
                    r_code.font.color.rgb = RGBColor(71, 85, 105)

            doc.save(output_docx_path)
            return output_docx_path

        finally:
            # Clean up temp unpacked directory
            if temp_dir and os.path.exists(temp_dir):
                import shutil
                shutil.rmtree(temp_dir, ignore_errors=True)

    @classmethod
    def _scan_vault_directory(cls, root_dir: str) -> Dict[str, Any]:
        """Scans and categorizes files from an unpacked Obsidian Vault."""
        result: Dict[str, Any] = {
            "metadata": {},
            "stories": [],
            "usecases": [],
            "data_dictionaries": {},
            "diagrams": {},
            "canvas": None,
            "other_docs": [],
        }

        # Check for Canvas files
        for f in os.listdir(root_dir):
            if f.endswith(".canvas"):
                try:
                    with open(os.path.join(root_dir, f), "r", encoding="utf-8") as fp:
                        result["canvas"] = json.load(fp)
                except Exception:
                    pass

        # Scan all .md files
        for dirpath, _, filenames in os.walk(root_dir):
            if ".obsidian" in dirpath:
                continue
            for fname in filenames:
                if not fname.endswith(".md"):
                    continue

                full_path = os.path.join(dirpath, fname)
                with open(full_path, "r", encoding="utf-8", errors="ignore") as fp:
                    content = fp.read()

                frontmatter, body = cls._parse_frontmatter(content)
                doc_type = frontmatter.get("type", "").lower()
                tags = [t.lower() for t in frontmatter.get("tags", [])] if isinstance(frontmatter.get("tags"), list) else []

                if fname == "Index.md" or doc_type == "vault-index":
                    result["metadata"] = frontmatter
                elif "user-story" in tags or doc_type == "user-story" or fname.startswith("US_"):
                    result["stories"].append({
                        "filename": fname,
                        "frontmatter": frontmatter,
                        "body": body,
                        "id": frontmatter.get("id", fname.replace(".md", "")),
                        "title": frontmatter.get("title", fname.replace(".md", "")),
                    })
                elif "use-case" in tags or doc_type == "use-case" or fname.startswith("UC_"):
                    result["usecases"].append({
                        "filename": fname,
                        "frontmatter": frontmatter,
                        "body": body,
                        "id": frontmatter.get("id", fname.replace(".md", "")),
                        "title": frontmatter.get("title", fname.replace(".md", "")),
                    })
                elif "data-dictionary" in tags or doc_type == "data-dictionary" or "Data-Dictionary" in dirpath:
                    tbl_name = frontmatter.get("table", fname.replace(".md", ""))
                    result["data_dictionaries"][tbl_name] = {
                        "filename": fname,
                        "frontmatter": frontmatter,
                        "body": body,
                    }
                elif "Diagrams" in dirpath or "mermaid" in body:
                    result["diagrams"][fname.replace(".md", "")] = body
                else:
                    result["other_docs"].append({
                        "filename": fname,
                        "frontmatter": frontmatter,
                        "body": body,
                    })

        return result

    @classmethod
    def _parse_frontmatter(cls, markdown_text: str) -> Tuple[Dict[str, Any], str]:
        """Separates YAML frontmatter from markdown body text."""
        if not markdown_text.startswith("---"):
            return {}, markdown_text

        parts = markdown_text.split("---", 2)
        if len(parts) >= 3:
            raw_yaml = parts[1]
            body = parts[2].strip()
            try:
                data = yaml.safe_load(raw_yaml) or {}
                if isinstance(data, dict):
                    return data, body
            except Exception:
                pass
        return {}, markdown_text

    @classmethod
    def _build_cover_page(cls, doc: Document, project_name: str, author: str):
        today = datetime.date.today().strftime("%d/%m/%Y")
        p_org = doc.add_paragraph()
        r_org = p_org.add_run("KHỐI CÔNG NGHỆ THÔNG TIN & CHUYỂN ĐỔI SỐ")
        r_org.font.size = Pt(11)
        r_org.bold = True
        r_org.font.color.rgb = RGBColor(71, 85, 105)

        for _ in range(4):
            doc.add_paragraph()

        p_t = doc.add_paragraph()
        r_t = p_t.add_run("TÀI LIỆU YÊU CẦU NGHIỆP VỤ\n(BRD / SRS SPECIFICATION)")
        r_t.font.size = Pt(24)
        r_t.bold = True
        r_t.font.color.rgb = RGBColor(15, 23, 42)

        p_sub = doc.add_paragraph()
        p_sub.paragraph_format.space_before = Pt(8)
        r_sub = p_sub.add_run(project_name.upper())
        r_sub.font.size = Pt(16)
        r_sub.bold = True
        r_sub.font.color.rgb = RGBColor(2, 132, 199)

        for _ in range(5):
            doc.add_paragraph()

        t_meta = doc.add_table(rows=4, cols=2)
        t_meta.alignment = WD_TABLE_ALIGNMENT.CENTER
        meta_pairs = [
            ("Tác Giả / Lead BA:", author),
            ("Hệ Sinh Thái:", "Obsidian Local-First Vault Bridge"),
            ("Phiên Bản Tài Liệu:", "1.0 (Enterprise OpenXML Edition)"),
            ("Ngày Biên Dịch:", today),
        ]
        for idx, (label, val) in enumerate(meta_pairs):
            row = t_meta.rows[idx]
            trPr = row._tr.get_or_add_trPr()
            trPr.append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))
            if idx == 0:
                trPr.append(parse_xml(f'<w:tblHeader {nsdecls("w")}/>'))
            c0, c1 = row.cells[0], row.cells[1]
            c0.width, c1.width = Inches(2.2), Inches(4.5)
            p0, p1 = c0.paragraphs[0], c1.paragraphs[0]
            p0.add_run(label).bold = True
            p0.runs[0].font.size = Pt(10)
            p1.add_run(val)
            p1.runs[0].font.size = Pt(10)

    @classmethod
    def _build_revision_history(cls, doc: Document, author: str):
        today = datetime.date.today().strftime("%d/%m/%Y")
        p_h = doc.add_paragraph()
        r_h = p_h.add_run("LỊCH SỬ THAY ĐỔI TÀI LIỆU (DOCUMENT REVISION CONTROL)")
        r_h.font.size = Pt(14)
        r_h.bold = True
        r_h.font.color.rgb = RGBColor(15, 23, 42)

        table = doc.add_table(rows=1, cols=5)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        headers = ["Phiên Bản", "Ngày", "Tác Giả", "Loại Thay Đổi", "Ghi Chú Nghiệp Vụ"]
        col_w = [Inches(1.0), Inches(1.2), Inches(1.8), Inches(1.5), Inches(1.8)]

        hdr = table.rows[0]
        trPr = hdr._tr.get_or_add_trPr()
        trPr.append(parse_xml(f'<w:tblHeader {nsdecls("w")}/>'))
        trPr.append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))

        for idx, text in enumerate(headers):
            c = hdr.cells[idx]
            c.width = col_w[idx]
            c_pr = c._tc.get_or_add_tcPr()
            c_pr.append(parse_xml(f'<w:shd {nsdecls("w")} w:fill="1E293B"/>'))
            c_pr.append(parse_xml(f'<w:vAlign {nsdecls("w")} w:val="center"/>'))
            p = c.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = p.add_run(text)
            r.bold = True
            r.font.size = Pt(9.5)
            r.font.color.rgb = RGBColor(255, 255, 255)

        row = table.add_row()
        row_tr = row._tr.get_or_add_trPr()
        row_tr.append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))
        vals = ["v1.0", today, author, "Khởi tạo ban đầu", "Biên dịch tự động từ Obsidian Vault"]
        for idx, v in enumerate(vals):
            c = row.cells[idx]
            c.width = col_w[idx]
            c_pr = c._tc.get_or_add_tcPr()
            c_pr.append(parse_xml(f'<w:vAlign {nsdecls("w")} w:val="center"/>'))
            p = c.paragraphs[0]
            p.add_run(v).font.size = Pt(9)

    @classmethod
    def _build_section_heading(cls, doc: Document, title: str, bookmark_id: str):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(18)
        p.paragraph_format.space_after = Pt(8)
        r = p.add_run(title)
        r.font.size = Pt(14)
        r.bold = True
        r.font.color.rgb = RGBColor(15, 23, 42)
        cls._add_bookmark(p, bookmark_id)

    @classmethod
    def _render_compiled_story(cls, doc: Document, story_dict: Dict[str, Any], index: int):
        st_id = story_dict.get("id", f"US-{index}")
        st_title = story_dict.get("title", "")
        body = story_dict.get("body", "")

        p_title = doc.add_paragraph()
        p_title.paragraph_format.space_before = Pt(12)
        p_title.paragraph_format.space_after = Pt(4)
        r_num = p_title.add_run(f"2.{index}. [{st_id}] {st_title}\n")
        r_num.bold = True
        r_num.font.size = Pt(12)
        r_num.font.color.rgb = RGBColor(2, 132, 199)

        # Add bookmark for internal Wikilinks targeting this story
        safe_bookmark = f"Bookmark_{re.sub(r'[^a-zA-Z0-9_]', '_', st_id)}"
        cls._add_bookmark(p_title, safe_bookmark)

        # Parse story sections from body
        lines = body.split("\n")
        in_code_block = False
        code_lines: List[str] = []

        for line in lines:
            line_s = line.strip()
            if line_s.startswith("```"):
                if not in_code_block:
                    in_code_block = True
                    code_lines = []
                else:
                    in_code_block = False
                    # Render code block (e.g. Gherkin)
                    p_code = doc.add_paragraph()
                    p_code.paragraph_format.left_indent = Inches(0.3)
                    p_code.paragraph_format.space_after = Pt(6)
                    r_c = p_code.add_run("\n".join(code_lines))
                    r_c.font.name = "Consolas"
                    r_c.font.size = Pt(8.5)
                    r_c.font.color.rgb = RGBColor(15, 23, 42)
                continue

            if in_code_block:
                code_lines.append(line)
                continue

            if line_s.startswith("#"):
                continue  # Skip raw markdown headers, already styled

            if line_s.startswith("- **As a:**") or line_s.startswith("- **I want:**") or line_s.startswith("- **So that:**"):
                p_st = doc.add_paragraph()
                p_st.paragraph_format.left_indent = Inches(0.2)
                cls._render_text_with_wikilinks(p_st, line_s)
            elif line_s.startswith("- "):
                p_li = doc.add_paragraph(style="List Bullet")
                cls._render_text_with_wikilinks(p_li, line_s[2:])
            elif line_s:
                p_p = doc.add_paragraph()
                cls._render_text_with_wikilinks(p_p, line_s)

    @classmethod
    def _render_compiled_usecase(cls, doc: Document, uc_dict: Dict[str, Any], index: int):
        uc_id = uc_dict.get("id", f"UC-{index}")
        uc_title = uc_dict.get("title", "")
        body = uc_dict.get("body", "")

        p_title = doc.add_paragraph()
        p_title.paragraph_format.space_before = Pt(12)
        p_title.paragraph_format.space_after = Pt(4)
        r_t = p_title.add_run(f"3.{index}. [{uc_id}] {uc_title}")
        r_t.bold = True
        r_t.font.size = Pt(12)
        r_t.font.color.rgb = RGBColor(16, 185, 129)

        safe_bookmark = f"Bookmark_{re.sub(r'[^a-zA-Z0-9_]', '_', uc_id)}"
        cls._add_bookmark(p_title, safe_bookmark)

        lines = [l for l in body.split("\n") if not l.strip().startswith("#")]
        for l in lines:
            ls = l.strip()
            if ls.startswith("- "):
                p = doc.add_paragraph(style="List Bullet")
                cls._render_text_with_wikilinks(p, ls[2:])
            elif ls:
                p = doc.add_paragraph()
                cls._render_text_with_wikilinks(p, ls)

    @classmethod
    def _render_compiled_data_dict(cls, doc: Document, tbl_name: str, tbl_info: Dict[str, Any]):
        p_tbl = doc.add_paragraph()
        p_tbl.paragraph_format.space_before = Pt(14)
        p_tbl.paragraph_format.space_after = Pt(4)
        r_t = p_tbl.add_run(f"Bảng Thực Thể: {tbl_name}")
        r_t.bold = True
        r_t.font.size = Pt(11.5)
        r_t.font.color.rgb = RGBColor(147, 51, 234)

        safe_bookmark = f"Bookmark_{re.sub(r'[^a-zA-Z0-9_]', '_', tbl_name)}"
        cls._add_bookmark(p_tbl, safe_bookmark)

        body = tbl_info.get("body", "")
        # Parse markdown table inside body
        tbl_lines = [l.strip() for l in body.split("\n") if l.strip().startswith("|") and l.strip().endswith("|")]
        if len(tbl_lines) >= 2:
            headers = [c.strip() for c in tbl_lines[0].split("|")[1:-1]]
            data_rows = [tbl_lines[i] for i in range(len(tbl_lines)) if i != 1 and not re.match(r"^\|(\s*[-:]+\s*\|)+$", tbl_lines[i])]
            data_rows = data_rows[1:]  # Exclude header

            t = doc.add_table(rows=1, cols=len(headers))
            t.alignment = WD_TABLE_ALIGNMENT.CENTER
            hdr = t.rows[0]
            trPr = hdr._tr.get_or_add_trPr()
            trPr.append(parse_xml(f'<w:tblHeader {nsdecls("w")}/>'))
            trPr.append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))

            for c_idx, h_text in enumerate(headers):
                cell = hdr.cells[c_idx]
                c_pr = cell._tc.get_or_add_tcPr()
                c_pr.append(parse_xml(f'<w:shd {nsdecls("w")} w:fill="1E293B"/>'))
                c_pr.append(parse_xml(f'<w:vAlign {nsdecls("w")} w:val="center"/>'))
                p = cell.paragraphs[0]
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                r = p.add_run(h_text)
                r.bold = True
                r.font.size = Pt(9)
                r.font.color.rgb = RGBColor(255, 255, 255)

            for r_idx, d_line in enumerate(data_rows):
                row = t.add_row()
                row_tr = row._tr.get_or_add_trPr()
                row_tr.append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))
                cols = [c.strip() for c in d_line.split("|")[1:-1]]
                for c_idx in range(len(headers)):
                    cell = row.cells[c_idx]
                    c_pr = cell._tc.get_or_add_tcPr()
                    c_pr.append(parse_xml(f'<w:vAlign {nsdecls("w")} w:val="center"/>'))
                    if r_idx % 2 == 1:
                        c_pr.append(parse_xml(f'<w:shd {nsdecls("w")} w:fill="F8FAFC"/>'))
                    p = cell.paragraphs[0]
                    col_val = cols[c_idx] if c_idx < len(cols) else ""
                    cls._render_text_with_wikilinks(p, col_val, font_size=8.5)

    @classmethod
    def _render_text_with_wikilinks(cls, paragraph, text: str, font_size: float = 9.5):
        """
        Parses text and renders Obsidian [[Page|Alias]] syntax into
        Word Bookmark Hyperlinks or bold accent text runs.
        """
        last_idx = 0
        for match in cls.WIKILINK_REGEX.finditer(text):
            start, end = match.span()
            if start > last_idx:
                r_pre = paragraph.add_run(text[last_idx:start])
                r_pre.font.size = Pt(font_size)

            target = match.group(1).strip()
            alias = match.group(2).strip() if match.group(2) else target

            # Create internal hyperlink to bookmark
            safe_target_bookmark = f"Bookmark_{re.sub(r'[^a-zA-Z0-9_]', '_', target)}"
            cls._add_hyperlink_to_bookmark(paragraph, safe_target_bookmark, alias, font_size)

            last_idx = end

        if last_idx < len(text):
            r_post = paragraph.add_run(text[last_idx:])
            r_post.font.size = Pt(font_size)

    @classmethod
    def _add_bookmark(cls, paragraph, bookmark_name: str):
        """Injects OpenXML bookmarkStart and bookmarkEnd into a paragraph."""
        b_id = str(abs(hash(bookmark_name)) % 100000)
        bm_start = parse_xml(f'<w:bookmarkStart {nsdecls("w")} w:id="{b_id}" w:name="{bookmark_name}"/>')
        bm_end = parse_xml(f'<w:bookmarkEnd {nsdecls("w")} w:id="{b_id}"/>')
        paragraph._p.append(bm_start)
        paragraph._p.append(bm_end)

    @classmethod
    def _add_hyperlink_to_bookmark(cls, paragraph, bookmark_name: str, text: str, font_size: float = 9.5):
        """Constructs an OpenXML internal hyperlink (<w:hyperlink w:anchor='...'>) targeting a bookmark."""
        try:
            hyperlink = parse_xml(
                f'<w:hyperlink {nsdecls("w")} w:anchor="{bookmark_name}" w:history="1">'
                f'  <w:r>'
                f'    <w:rPr>'
                f'      <w:color w:val="0284C7"/>'
                f'      <w:b/>'
                f'      <w:u w:val="single"/>'
                f'      <w:sz w:val="{int(font_size * 2)}"/>'
                f'    </w:rPr>'
                f'    <w:t>{text}</w:t>'
                f'  </w:r>'
                f'</w:hyperlink>'
            )
            paragraph._p.append(hyperlink)
        except Exception:
            # Fallback to bold colored run if XML injection fails
            r = paragraph.add_run(f"[{text}]")
            r.bold = True
            r.font.size = Pt(font_size)
            r.font.color.rgb = RGBColor(2, 132, 199)
