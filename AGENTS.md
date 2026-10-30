# AGENTS.md — Guidance for Antigravity AI Agents

Welcome, AI Agent! This file is your operational manual for `antigravity-doc-handler`.

## 📦 What is this Repository?

`antigravity-doc-handler` is a specialized Python toolkit designed for:
1. **High-Fidelity Document Processing**: Bidirectional conversion between PDF, DOCX, and Markdown with zero data loss, OpenXML table repair, and decoupled styling.
2. **Precision Technical Diagram Engine**: Programmatic and declarative generation of publication-grade Hub & Spoke architecture and mobile screen flow diagrams (300+ DPI, vector SVG + headless Chromium PNG).

---

## 🛠 Available Skills & Rules

The workspace includes preconfigured Antigravity customizations:
- **Skill 1**: [doc-handler](.agents/skills/doc-handler/SKILL.md) — PDF/Word/Markdown conversion, OpenXML table repair, document inspection, and JSON supremacy mental model.
- **Skill 2**: [excel-automation](.agents/skills/excel-automation/SKILL.md) — Template-driven OpenXML Excel processing, token extraction, live KPI linking, 6 enterprise subsystems, and streaming mode.
- **Skill 3**: [felo-slides](.agents/skills/felo-slides/SKILL.md) — 16:9 widescreen presentation synthesis, Slide Master inheritance, and locked aspect-ratio diagram injection.
- **Skill 4**: [mxgraph-diagram-engineering](.agents/skills/mxgraph-diagram-engineering/SKILL.md) — Draw.io (.drawio / XML) pure mxGraphModel generation, layout, orthogonal routing, and monochrome line-art.
- **Skill 5**: [technical-diagrams](.agents/skills/technical-diagrams/SKILL.md) — Declarative diagram rendering (Canvas, Mermaid, PlantUML), orthogonal Manhattan routing, and collision prevention.
- **Skill 6**: [security-audit](.agents/skills/security-audit/SKILL.md) — Cloudflare-seeded defensive security audit harness (Reconnaissance, Coverage-Led Hunting, Adversarial Validation, Structured Schema, Zero-Dependency Verification, Reporting).
- **Rule 1**: [rule_technical_diagram_standards.md](.agents/rules/rule_technical_diagram_standards.md) — English screen names, codebase-truth action labels, multi-line wrapping (`\n`), and automated AABB collision detection.
- **Rule 2**: [rule_decoupled_document_converter.md](.agents/rules/rule_decoupled_document_converter.md) — Table OpenXML invariants (`cantSplit`, `tblHeader`, `vAlign="center"`) and decoupled styling.
- **Rule 3**: [rule_feature_lifecycle_and_doc_sync.md](.agents/rules/rule_feature_lifecycle_and_doc_sync.md) — Documentation synchronization (`CHANGELOG.md`, `ARCHITECTURE.md`, `README.md`, and Core Docs Triad) on feature completions (`ERR_DOC_SYNC_001`).
- **Rule 4**: [rule_version_synchronization_and_release_invariants.md](.agents/rules/rule_version_synchronization_and_release_invariants.md) — 10-anchor synchronization and automated build on version bumps (`ERR_VER_001`).
- **Rule 5**: [rule_database_erd_standards.md](.agents/rules/rule_database_erd_standards.md) — Crow's foot notation, table sizing, and relational database ERD invariants.
- **Rule 6**: [rule_excel_template_preservation_and_ux.md](.agents/rules/rule_excel_template_preservation_and_ux.md) — Invariants E1–E14, token extraction, and live dynamic formulas.
- **Rule 7**: [rule_enterprise_document_and_spreadsheet_qa.md](.agents/rules/rule_enterprise_document_and_spreadsheet_qa.md) — Document & spreadsheet QA invariants (`ERR_XLSX_001`–`014`, `ERR_DOCX_001`–`005`).
- **Rule 8**: [rule_mxgraph_drawio_diagram_standards.md](.agents/rules/rule_mxgraph_drawio_diagram_standards.md) — Draw.io mxGraphModel invariants (`MX_INV_01`–`20`).
- **Rule 9**: [rule_learning_and_skill_authoring.md](.agents/rules/rule_learning_and_skill_authoring.md) — Meta-rule for authoring rules and skills, auto-updating `README.md` and `AGENTS.md`.
- **Rule 10**: [rule_context_recovery_protocol.md](.agents/rules/rule_context_recovery_protocol.md) — Recovery invariants and checkpoint protocols.
- **Rule 11**: [rule_quality_gate_and_verification.md](.agents/rules/rule_quality_gate_and_verification.md) — Mandatory multi-layer quality gate audit before task completion.
- **Rule 12**: [rule_security_audit_standards.md](.agents/rules/rule_security_audit_standards.md) — Mandatory security principles, Adversarial Verifier protocol, and document safety invariants (`ERR_SEC_001`–`005`).

---

## 💻 CLI Commands Cheat Sheet

When user requests you to perform document tasks or draw diagrams, run these commands:

| Task | Command Line |
|---|---|
| **Zero-Setup Environment Doctor** | `python ai_tools_cli.py doctor [--json]` |
| **Model Context Protocol (MCP) Server** | `python mcp_server.py` (STDIO JSON-RPC 2.0) |
| **Web Studio Server & SSE Queue** | `python web_api_server.py 8000` |
| **Unified Multi-Engine Diagram Render** | `python ai_tools_cli.py diagram-render <spec.json> -o <out.png>` |
| **Render PlantUML Diagram** | `python ai_tools_cli.py plantuml-render <spec.json> -o <out.png> --dpi 300` |
| **Render Mermaid Diagram** | `python ai_tools_cli.py mermaid-render <spec.json> -o <out.png>` |
| **Render SVG Canvas Precision Diagram** | `python ai_tools_cli.py spec-render <spec.json> -o <out.png> --scale 3` |
| **Direct Diagram Engine CLI** | `python spec_diagram_engine.py --spec <spec.json> --out <out.png> --scale 3` |
| **Interactive Canvas Diagram Editor** | `python diagram_editor.py --spec <spec.json>` |
| **Convert PDF to DOCX** | `python -m ai_tools_cli convert <input.pdf> -o <output.docx>` |
| **Convert DOCX to Decoupled MD** | `python -m ai_tools_cli convert <document.docx> -t md` |
| **Convert MD to Styled DOCX** | `python -m ai_tools_cli convert <document.md> -t docx --style <doc.style.yaml>` |
| **Inspect DOCX Headings & Tables** | `python -m ai_tools_cli inspect-doc <document.docx>` |
| **Insert Image into DOCX** | `python -m ai_tools_cli insert-diagram <doc.docx> <img.png> -s "Heading" -c "Caption"` |
| **Legal Redline Diff (Track Changes)** | `python ai_tools_cli.py docx-diff <v1.docx> <v2.docx> -o <diff.docx>` |
| **Batch Mail Merge & Templating** | `python ai_tools_cli.py docx-mail-merge <tpl.docx> <data.json> -o <merged.docx>` |
| **Unified QA Audit (13 Gates / OpenXML)** | `python ai_tools_cli.py qa-audit <doc.docx|sheet.xlsx> [--security] [--sanitize] [--strict]` |
| **Document Security Audit & Sanitizer** | `python ai_tools_cli.py doc-security <doc> [--sanitize] [-o clean_doc]` |
| **Vietnamese OCR & Scan Recovery** | `python ai_tools_cli.py ocr-scan <scan.pdf> -o <out.md> [--lang vie+eng]` |
| **Build 16:9 Presentation from Spec** | `python ai_tools_cli.py pptx-build <spec.json> -o <deck.pptx> [--template <tpl.pptx>]` |
| **Export Slide PNG Previews for QA** | `python ai_tools_cli.py pptx-preview <deck.pptx> -o <preview_dir> [--slides 1,5,11]` |
| **PDF Pro Studio Suite (16 Actions)** | `python ai_tools_cli.py pdf-tools {images-to-pdf,export-images-zip,sign,flatten,reorganize,redact,protect,giap-lai,...}` |
| **Images to PDF** | `python ai_tools_cli.py pdf-tools images-to-pdf --inputs <img1> <img2> -o <out.pdf>` |
| **PDF to 300 DPI Images ZIP** | `python ai_tools_cli.py pdf-tools export-images-zip -i <doc.pdf> -o <bundle.zip> --dpi 300` |
| **Sign PDF with Stamp/Signature** | `python ai_tools_cli.py pdf-tools sign -i <doc.pdf> -o <signed.pdf> --stamp <sig.png> --page 1` |
| **Flatten & Sanitize PDF** | `python ai_tools_cli.py pdf-tools flatten -i <doc.pdf> -o <flat.pdf> [--strict]` |
| **Reorganize Pages with Blank Insertion** | `python ai_tools_cli.py pdf-tools reorganize -i <doc.pdf> -o <reorg.pdf> --plan "1,blank,2"` |
| **Set Excel Cell Value/Formula** | `python ai_tools_cli.py xlsx-set <file.xlsx> --sheet <sheet> --cell <cell> --value <val>` |
| **Insert Excel Rows & Shift Formulas** | `python ai_tools_cli.py xlsx-insert-rows <file.xlsx> --sheet <sheet> --start-row <row> --count <n>` |
| **Inspect Word Template Tokens** | `python ai_tools_cli.py docx-tokens <tpl.docx> [--json]` |
| **Atomic Word Text Replace** | `python ai_tools_cli.py docx-replace <doc.docx> --data <patch.json>` |
| **Sync Core Docs Triad (MD->DOCX+PDF)** | `python tools/sync_core_docs.py [--build|--check]` |
| **Launch Desktop GUI** | `python main.py` |

---

## 🗺️ Schema-First Routing Table (AI Agent Dispatch)

When executing tasks programmatically via Python, MCP, or CLI, dispatch according to this matrix:

| User Task / Intent | Input Type | MCP Tool | CLI Subcommand | Core Module |
|---|---|---|---|---|
| **System Diagnostics** | Environment | `env_doctor()` | `ai_tools_cli.py doctor` | `core.env_doctor.run_doctor` |
| **Convert Doc (Lossless)** | PDF / DOCX / MD / HTML | `doc_convert(...)` | `ai_tools_cli.py convert` | `core.converter.engine.convert_universal` |
| **PDF Pro Suite & Security** | PDF / Images / Stamp | `pdf_quick_tools(...)` | `ai_tools_cli.py pdf-tools` | `core.converter.pdf_toolkit` |
| **Inspect Document AST** | DOCX / PPTX / XLSX | `doc_inspect(...)` | `ai_tools_cli.py docx-read` / `pptx-inspect` | `core.docx.reader`, `core.pptx.reader` |
| **Legal Redline Diff (Track Changes)** | DOCX v1 + DOCX v2 | `doc_redline_diff(...)` | `ai_tools_cli.py docx-diff` | `core.docx.diff_engine.compare_docx_documents` |
| **Batch Mail Merge & Templating** | DOCX + JSON / CSV | `doc_mail_merge(...)` | `ai_tools_cli.py docx-mail-merge` | `core.docx.writer` |
| **Unified Document QA Audit** | DOCX / XLSX | `doc_qa_audit(...)` | `ai_tools_cli.py qa-audit` | `core.qa.unified_audit` |
| **Vietnamese OCR & Scan Recovery** | Scanned PDF / Image | `doc_ocr_scan(...)` | `ai_tools_cli.py ocr-scan` | `core.converter.ocr_pipeline.VietnameseOcrPipeline` |
| **Inject Diagram Image** | DOCX / PPTX + PNG | `doc_insert_diagram(...)` | `ai_tools_cli.py docx-inject` / `pptx-insert-diagram` | `core.docx.writer`, `core.pptx.engine` |
| **Render Technical Diagram** | Spec JSON / Raw Code | `diagram_render(...)` | `ai_tools_cli.py diagram-render` | `core.diagrams.bridge` |
| **Synthesize 16:9 Presentation** | DOCX / PDF / MD | `pptx_synthesize(...)` | `ai_tools_cli.py docx-to-pptx` / `pptx-from-md` | `core.pptx.docx_to_pptx`, `core.pptx.engine` |
| **Excel Template Mutation** | XLSX + Patch JSON | (via REST API) | `ai_tools_cli.py xlsx-mutate` | `core.xlsx.writer.mutate_template_excel` |
| **Excel 12-Gate Audit** | XLSX Target + Ref | (via REST API) | `ai_tools_cli.py xlsx-validate` | `core.xlsx.validator.validate_excel_sheet` |
| **Atomic Excel Cell Edit** | XLSX + Cell + Value | `xlsx_set_cell(...)` / `xlsx_set_range(...)` | `ai_tools_cli.py xlsx-set` | `core.xlsx.writer.UniversalExcelEngine` |
| **Atomic Excel Row Expansion** | XLSX + Row + Count | `xlsx_insert_rows(...)` | `ai_tools_cli.py xlsx-insert-rows` | `core.xlsx.writer.UniversalExcelEngine` |
| **Inspect Word Placeholders** | DOCX Template | `docx_inspect_placeholders(...)` | `ai_tools_cli.py docx-tokens` | `core.docx.mail_merge.inspect_template_tokens` |
| **Atomic Word Text Patching** | DOCX + Replacements | `docx_replace_text(...)` | `ai_tools_cli.py docx-replace` | `core.docx.writer.patch_docx_template` |
| **Atomic Word Paragraph/Table** | DOCX + Content | `docx_insert_paragraph(...)` / `docx_insert_table(...)` | (via MCP Tool) | `core.docx.writer` |
| **Document Security Audit & Sanitizer** | Office Doc / CSV / Zip | `doc_security_audit(...)` | `ai_tools_cli.py doc-security` | `core.security.document_auditor`, `core.security.sanitizer` |
| **PKCS#7 / PAdES Signing & PDF/A-2b** | PDF + Optional PFX | `pdf_crypto_sign_and_pdfa(...)` | `ai_tools_cli.py pdf-tools {crypto-sign,pdfa}` | `core.converter.pdf_toolkit` |
| **Excel-to-Brief Report (Word/MD/PDF)** | XLSX Workbook | `xlsx_to_brief_report(...)` | `ai_tools_cli.py xlsx-to-brief` | `core.xlsx.reporting.excel_to_brief` |
| **Cross-Workbook Consolidation** | Multiple XLSX | `xlsx_consolidate_workbooks(...)` | `ai_tools_cli.py xlsx-consolidate` | `core.xlsx.reporting.consolidate_workbooks` |
| **Declarative Workflow Pipeline** | File/Dir + Spec | `doc_run_pipeline(...)` | `ai_tools_cli.py run-pipeline` | `core.pipeline.engine` |
| **AST Code-to-Diagram & SVG Export** | Python / SQL / Prisma | `diagram_from_code(...)` | `ai_tools_cli.py code-to-diagram` | `core.diagrams.code_to_diagram` |
| **Launch Native Web Studio App Mode** | Desktop App | (CLI / GUI Button) | `ai_tools_cli.py desktop-web` | `ui.webview_launcher.launch_native_web_studio` |

---

## ⚡ Core Engine Architecture

```
antigravity-doc-handler/
├── core/                       # Layer 3: Core Document & Diagram Engines
│   ├── converter/              # Universal Converter, Decoupled MD, OCR, PDF Toolkit
│   ├── diagrams/               # Diagram Bridge, mxGraph, Code-to-Diagram, Editor, Legacy
│   ├── docx/                   # OpenXML Reader/Writer, Lists, Math, Diff, Mail Merge, Presets
│   ├── pptx/                   # 16:9 Widescreen Engine, DOCX-to-PPTX, Charts, Math
│   ├── xlsx/                   # Universal Excel Engine, Validator, Streaming, Formula Playbook, Reporting
│   ├── security/               # Cloudflare Document Auditor & Deep CDR / PII Sanitizer
│   ├── pipeline/               # Declarative Multi-Step Workflow Pipeline Engine
│   └── env_doctor.py           # 12-Component Environment Health Doctor
├── server/                     # Layer 2: Web Studio REST API (:8000), MCP Server (28 Tools), SSE Jobs, SQLite
├── cli/                        # Layer 2: Unified AI CLI Dispatcher & Subcommands
├── ui/                         # Layer 1: Desktop GUI (Tkinter) & Native WebView App Launcher
├── src/                        # Layer 1: Vite ES6 Web Studio Frontend & Modular Controllers
├── tools/                      # QA Diagnostics, Version Sync, Core Docs Sync & plantuml.jar
│   ├── unified_qa_diagnostic.py # 13-Gate Unified QA Diagnostic & Self-Repair Engine
│   ├── bump_version.py         # 10-Anchor Version Synchronization Tool (ERR_VER_001)
│   ├── sync_core_docs.py       # Core Docs Triad Compiler & Semantic Auditor (ERR_DOC_SYNC_001)
│   └── plantuml.jar            # Local PlantUML JAR (v1.2026+, bundled C4 standard library)
├── specs/                      # Declarative Diagram Specs (flow_*.json, master_erd_*.json)
├── tests/                      # 323 Automated Test Suites (tests/test_*.py)
├── main.py                     # Official Entrypoint: Desktop GUI & Native Web App Mode
├── ai_tools_cli.py             # Official Entrypoint: Unified CLI Dispatcher for AI Agents
├── web_api_server.py           # Official Entrypoint: Web Studio REST API Server
├── mcp_server.py               # Official Entrypoint: Model Context Protocol STDIO Server
├── spec_diagram_engine.py      # Official Entrypoint: SVG Canvas Precision Diagram CLI
├── diagram_editor.py           # Official Entrypoint: Interactive Canvas Diagram Editor
└── .agents/                    # Antigravity Customizations (6 Skills & 12 Rules)
```

---

## 🎯 Important Invariants & Diagram Selection Rules

### Diagram Engine Selection Rules (MANDATORY):
1. **Always Read `"engine"` Field First**: Dispatch to `"canvas"`, `"mermaid"`, or `"plantuml"`.
2. **Tuân Thủ Chuẩn Phân Loại 12 Loại Sơ Đồ Kỹ Thuật**:
   - **`canvas` Engine** (1):
     - Screen Flow (Interactive) — Tọa độ pixel-perfect, kéo thả trên `diagram_editor.py`.
   - **`mermaid` Engine** (5):
     - ERD (Database Schema) — Nền phẳng pastel hiện đại, gọn gàng cho phân hệ 3–15 bảng.
     - Sequence Diagram — Chuỗi gọi API thanh thoát, đánh số tự động `autonumber`.
     - Flowchart / Process — Bẻ nhánh if/else tự do, đổi màu khối nhanh bằng CSS/style.
     - State Diagram — Trạng thái bo góc tròn, màu sắc hiện đại.
     - Mind Map — Phân rã tính năng nhanh, màu pastel chia nhánh trực quan.
   - **`plantuml` Engine** (6):
     - Use Case Diagram — Actor người que, quan hệ `<<include>>`, `<<extend>>` chuẩn UML.
     - C4 Architecture — Thư viện C4 chuẩn quốc tế (`<C4/C4_Context>`), không vỡ dây container.
     - Component / Deployment — Đúng hình khối 3D `node`, `database`, `component`.
     - Activity Swimlane — Phân làn cột (`|Partition|`) thẳng đứng tuyệt đối, không đè dây.
     - Class Diagram — Hỗ trợ trọn vẹn OOP (`+`, `-`, `#`, generics `<T>`, composition).
     - Package Diagram — Hỗ trợ stereotype `<<Folder>>` trực quan cho cấu trúc thư mục.
3. **Tuân thủ 4 Bẫy Kỹ Thuật Windows Cho PlantUML**:
   - C4: Luôn dùng `!include <C4/C4_Context>` Standard Library nội bộ, tuyệt đối KHÔNG dùng URL Raw GitHub.
   - Font: Ghi file UTF-8 và truyền cờ JVM `-charset UTF-8` để tiếng Việt không bị vỡ.
   - Limit: Luôn truyền `-DPLANTUML_LIMIT_SIZE=16384` để ảnh kiến trúc/sơ đồ lớn không bị mờ hoặc cắt cụt.
   - Auto-Inject: Luôn kiểm tra `inject_into` để tự động nhúng vào DOCX sau khi render PNG.
4. **Table Invariants**: Whenever modifying Word `.docx` tables, ensure `<w:cantSplit/>`, `<w:tblHeader/>`, and `<w:vAlign w:val="center"/>` are present.
5. **Diagram Aspect Ratio**: Keep technical diagrams within the $1.6:1 - 1.85:1$ aspect ratio (e.g. $1400 \times 770\text{px}$ or $1360 \times 720\text{px}$) to perfectly fit standard portrait A4 margins ($14\text{cm}$ print width) without font shrinkage.
6. **JSON Spec Supremacy**: Always author declarative `.json` specs for diagrams. Do not rely on uncontrolled auto-layout.
7. **Standardized English Screen Names**: All node box titles must be 100% technical English (`Login Screen`, `Home Dashboard`, `Cart Screen`). Never put Vietnamese inside node boxes.
8. **Codebase-Truth Action Phrasing**: Use English verbs (`Click`, `Tap`, `Select`) + quoted original button labels (`Click "Đăng nhập"`). Default to 100% English for multi-language projects.
9. **Multi-line Wrapping & Collision Avoidance**: Always break action labels across multiple lines with `\n` when length $> 18-22$ characters. Pre-validate using engine AABB collision detection (`[WARN] Label collision detected`).
