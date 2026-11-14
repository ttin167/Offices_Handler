# AGENTS.md — Guidance for Antigravity AI Agents

Welcome, AI Agent! This file is your operational manual for `antigravity-doc-handler`.

## 📦 What is this Repository?

`antigravity-doc-handler` is a specialized Python toolkit designed for:

1. **High-Fidelity Document Processing**: Bidirectional conversion between PDF, DOCX, and Markdown with zero data loss, OpenXML table repair, and decoupled styling.
2. **Precision Technical Diagram Engine**: Programmatic and declarative generation of publication-grade Hub & Spoke architecture and mobile screen flow diagrams (300+ DPI, vector SVG + headless Chromium PNG).
3. **Pure Office IT Automation & AI Copilot Workstation**: 6-Tool Excel Power Suite, Office Work Tracker, 57-Formula Playbook & AST Evaluator, Vietnamese Word Mastery (Decree 30/2020/NĐ-CP & `vn_words`), PDF Pro 17-Tool Suite, Universal FTS5 RAG & Hybrid AI Copilot.

---

## 🗺️ Schema-First Routing Table (AI Agent Direct Dispatch)

> [!IMPORTANT]
> **CRITICAL DIRECTIVE FOR AI AGENTS**:
> Before performing any directory scan, recursive grep, or guessing file locations, **YOU MUST** consult the Routing Table below and navigate directly to the target file path. **DO NOT scan directories.**

| User Task / Intent | Input Type | MCP Tool | CLI Subcommand | Target File Path (Core Implementation) |
| --- | --- | --- | --- | --- |
| **System Diagnostics** | Environment | `env_doctor()` | `ai_tools_cli.py doctor` | `core/env_doctor.py` (`run_doctor`) |
| **Convert Doc (Lossless)** | PDF / DOCX / MD / HTML | `doc_convert(...)` | `ai_tools_cli.py convert` | `core/converter/engine.py` (`convert_universal`) |
| **PDF Pro Suite & Security** | PDF / Images / Stamp | `pdf_quick_tools(...)` | `ai_tools_cli.py pdf-tools` | `core/converter/pdf_toolkit.py` (`PDFQuickToolkit`) |
| **Inspect Document AST** | DOCX / PPTX / XLSX | `doc_inspect(...)` | `ai_tools_cli.py docx-read` / `pptx-inspect` | `core/docx/reader.py`, `core/pptx/reader.py` |
| **Legal Redline Diff (Track Changes)** | DOCX v1 + DOCX v2 | `doc_redline_diff(...)` | `ai_tools_cli.py docx-diff` | `core/docx/diff_engine.py` (`compare_docx_documents`) |
| **Batch Mail Merge & Templating** | DOCX + JSON / CSV | `doc_mail_merge(...)` | `ai_tools_cli.py docx-mail-merge` | `core/docx/writer.py`, `core/docx/mail_merge.py` |
| **Unified Document QA Audit** | DOCX / XLSX | `doc_qa_audit(...)` | `ai_tools_cli.py qa-audit` | `core/qa/unified_audit.py` (`run_unified_qa`) |
| **Vietnamese OCR & Scan Recovery** | Scanned PDF / Image | `doc_ocr_scan(...)` | `ai_tools_cli.py ocr-scan` | `core/converter/ocr_pipeline.py` (`VietnameseOcrPipeline`) |
| **Inject Diagram Image** | DOCX / PPTX + PNG | `doc_insert_diagram(...)` | `ai_tools_cli.py docx-inject` / `pptx-insert-diagram` | `core/docx/writer.py` (`inject_diagram_into_docx`), `core/pptx/engine.py` |
| **Render Technical Diagram** | Spec JSON / Raw Code | `diagram_render(...)` | `ai_tools_cli.py diagram-render` | `core/diagrams/bridge.py` (`render_diagram_file_to_png`) |
| **Synthesize 16:9 Presentation** | DOCX / PDF / MD | `pptx_synthesize(...)` | `ai_tools_cli.py docx-to-pptx` / `pptx-from-md` | `core/pptx/docx_to_pptx.py` (`convert_docx_to_pptx`), `core/pptx/engine.py` |
| **Excel Template Mutation** | XLSX + Patch JSON | (via REST API) | `ai_tools_cli.py xlsx-mutate` | `core/xlsx/writer.py` (`mutate_template_excel`) |
| **Excel 12-Gate Audit** | XLSX Target + Ref | (via REST API) | `ai_tools_cli.py xlsx-validate` | `core/xlsx/validator.py` (`validate_excel_sheet`) |
| **Atomic Excel Cell Edit** | XLSX + Cell + Value | `xlsx_set_cell(...)` / `xlsx_set_range(...)` | `ai_tools_cli.py xlsx-set` | `core/xlsx/writer.py` (`UniversalExcelEngine.set_cell`, `set_range`) |
| **Atomic Excel Row Expansion** | XLSX + Row + Count | `xlsx_insert_rows(...)` | `ai_tools_cli.py xlsx-insert-rows` | `core/xlsx/writer.py` (`UniversalExcelEngine.insert_rows`) |
| **Inspect Word Placeholders** | DOCX Template | `docx_inspect_placeholders(...)` | `ai_tools_cli.py docx-tokens` | `core/docx/mail_merge.py` (`inspect_template_tokens`) |
| **Atomic Word Text Patching** | DOCX + Replacements | `docx_replace_text(...)` | `ai_tools_cli.py docx-replace` | `core/docx/writer.py` (`patch_docx_template`) |
| **Atomic Word Paragraph/Table** | DOCX + Content | `docx_insert_paragraph(...)` / `docx_insert_table(...)` | (via MCP Tool) | `core/docx/writer.py` (`insert_paragraph_at_anchor`, `insert_table_at_anchor`) |
| **Document Security Audit & Sanitizer** | Office Doc / CSV / Zip | `doc_security_audit(...)` | `ai_tools_cli.py doc-security` | `core/security/document_auditor.py`, `core/security/sanitizer.py` |
| **PKCS#7 / PAdES Signing & PDF/A-2b** | PDF + Optional PFX | `pdf_crypto_sign_and_pdfa(...)` | `ai_tools_cli.py pdf-tools {crypto-sign,pdfa}` | `core/converter/pdf_toolkit.py` (`crypto_sign_pdf`, `convert_to_pdfa`) |
| **Excel-to-Brief Report (Word/MD/PDF)** | XLSX Workbook | `xlsx_to_brief_report(...)` | `ai_tools_cli.py xlsx-to-brief` | `core/xlsx/reporting.py` (`excel_to_brief`) |
| **Cross-Workbook Consolidation** | Multiple XLSX | `xlsx_consolidate_workbooks(...)` | `ai_tools_cli.py xlsx-consolidate` | `core/xlsx/reporting.py` (`consolidate_workbooks`) |
| **Declarative Workflow Pipeline** | File/Dir + Spec | `doc_run_pipeline(...)` | `ai_tools_cli.py run-pipeline` | `core/pipeline/engine.py` (`WorkflowPipelineEngine`) |
| **AST Code-to-Diagram & SVG Export** | Python / SQL / Prisma | `diagram_from_code(...)` | `ai_tools_cli.py code-to-diagram` | `core/diagrams/code_to_diagram.py` (`generate_diagram_from_code`) |
| **Launch Native Web Studio App Mode** | Desktop App | (CLI / GUI Button) | `ai_tools_cli.py desktop-web` | `ui/webview_launcher.py` (`launch_native_web_studio`) |
| **Excel Power Clean (Trim/Date/CCCD/Dedup)** | XLSX Workbook | `xlsx_power_clean(...)` | `ai_tools_cli.py xlsx-power-clean` | `core/xlsx/power_tools.py` (`clean_and_standardize_workbook`) |
| **Excel Split / Merge Columns (VN Names)** | XLSX + Columns | `xlsx_transform_columns(...)` | `ai_tools_cli.py xlsx-transform-cols` | `core/xlsx/power_tools.py` (`transform_workbook_columns`) |
| **Excel Pivot Summary & Native Chart** | XLSX Workbook | `xlsx_pivot_chart(...)` | `ai_tools_cli.py xlsx-pivot-chart` | `core/xlsx/power_tools.py` (`build_pivot_summary_and_chart`) |
| **Excel No-Code XLOOKUP Join** | XLSX v1 (+ XLSX v2) | `xlsx_xlookup_join(...)` | `ai_tools_cli.py xlsx-xlookup-join` | `core/xlsx/power_tools.py` (`xlookup_join_workbooks`) |
| **Excel Split Workbook by Column (.zip)** | XLSX + Split Col | `xlsx_split_by_column(...)` | `ai_tools_cli.py xlsx-split-col` | `core/xlsx/power_tools.py` (`split_workbook_by_column`) |
| **Excel Visual Cell Diff & _Diff_Summary** | XLSX v1 + XLSX v2 | `xlsx_diff_visual(...)` | `ai_tools_cli.py xlsx-diff` | `core/xlsx/power_tools.py` (`diff_workbooks_visual`) |
| **Office Work & Daily Task Tracker** | Tracking Workbooks | (via REST API `/api/xlsx/tracker/*`) | `ai_tools_cli.py xlsx-track` | `core/xlsx/tracker.py` (`OfficeWorkTracker`) |
| **Excel 57-Formula Playbook & Demo** | Formula Catalog | `xlsx_formula_helper(...)` | `ai_tools_cli.py xlsx-formulas` | `core/xlsx/formula_playbook.py` (`FORMULA_CATALOG`) |
| **Word Mastery Polish (NĐ30 & Page X/Y)** | DOCX Document | `docx_polish(...)` | `ai_tools_cli.py docx-polish` | `core/docx/toolkit.py` (`polish_docx_document`) |
| **Word Multi-File Merge & Media Remap** | Multiple DOCX | `docx_merge(...)` | `ai_tools_cli.py docx-merge` | `core/docx/toolkit.py` (`merge_docx_files`) |
| **Word Split by Heading Level (.zip)** | DOCX Document | `docx_split(...)` | `ai_tools_cli.py docx-split` | `core/docx/toolkit.py` (`split_docx_by_heading`) |
| **Word Watermark & Header/Footer Brand** | DOCX + Text | `docx_watermark_hf(...)` | `ai_tools_cli.py docx-watermark` | `core/docx/toolkit.py` (`apply_docx_watermark_and_hf`) |
| **VN Amount in Words (`vn_words`) & Tokens** | Number / Date / Spec | `doc_mail_merge(...)` | `ai_tools_cli.py docx-mail-merge` | `core/docx/design_system.py` (`number_to_vietnamese_words`) |
| **List Hardware System Printers** | Hardware Spooler | `doc_list_printers(...)` | `ai_tools_cli.py print-list` | `core/printing/printer_manager.py` (`list_printers`) |
| **Inspect Printer Status & Spooler Queue** | Printer Name | `doc_printer_status(...)` | `ai_tools_cli.py print-status` | `core/printing/printer_manager.py` (`get_printer_status`) |
| **Enterprise Silent Print Document** | PDF / DOCX + Print Spec | `doc_print_job(...)` | `ai_tools_cli.py print-doc` | `core/printing/printer_manager.py` (`print_document`) |
| **Offline AST Formula Evaluator** | XLSX Workbook | `xlsx_evaluate_formulas(...)` | `ai_tools_cli.py xlsx-eval` | `core/xlsx/formula_evaluator.py` (`ASTFormulaEvaluator`) |
| **Universal FTS5 RAG Search** | Multi-Format File + Query | `universal_rag_search(...)` | `ai_tools_cli.py rag-search` | `core/copilot/universal_rag.py` (`UniversalRAGEngine`) |
| **Hybrid AI Copilot Assistant** | Prompt + RAG Context | `copilot_assist(...)` | `ai_tools_cli.py copilot-chat` | `core/copilot/copilot_provider.py` (`CopilotOrchestrator`) |
| **Periodic System Maintenance (7-Pillars)** | System Workspace | `system_maintenance(...)` | `ai_tools_cli.py system-maintenance` | `core/system_maintenance.py` (`run_periodic_maintenance`) |

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
| --- | --- |
| **Zero-Setup Environment Doctor** | `python ai_tools_cli.py doctor [--json]` |
| **Periodic System Maintenance (7-Pillars)** | `python ai_tools_cli.py system-maintenance [--dry-run] [--json]` |
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
| **Unified QA Audit (13 Gates / OpenXML)** | `python ai_tools_cli.py qa-audit <doc.docx | sheet.xlsx> [--security] [--sanitize] [--strict]` |
| **Document Security Audit & Sanitizer** | `python ai_tools_cli.py doc-security <doc> [--sanitize] [-o clean_doc]` |
| **Vietnamese OCR & Scan Recovery** | `python ai_tools_cli.py ocr-scan <scan.pdf> -o <out.md> [--lang vie+eng]` |
| **Build 16:9 Presentation from Spec** | `python ai_tools_cli.py pptx-build <spec.json> -o <deck.pptx> [--template <tpl.pptx>]` |
| **Export Slide PNG Previews for QA** | `python ai_tools_cli.py pptx-preview <deck.pptx> -o <preview_dir> [--slides 1,5,11]` |
| **PDF Pro Studio Suite (17 Actions)** | `python ai_tools_cli.py pdf-tools {merge,split,compress,watermark,images-to-pdf,export-images-zip,sign,flatten,reorganize,redact,protect,giap-lai,crypto-sign,pdfa,...}` |
| **Images to PDF** | `python ai_tools_cli.py pdf-tools images-to-pdf --inputs <img1> <img2> -o <out.pdf>` |
| **PDF to 300 DPI Images ZIP** | `python ai_tools_cli.py pdf-tools export-images-zip -i <doc.pdf> -o <bundle.zip> --dpi 300` |
| **Sign PDF with Stamp/Signature** | `python ai_tools_cli.py pdf-tools sign -i <doc.pdf> -o <signed.pdf> --stamp <sig.png> --page 1` |
| **Flatten & Sanitize PDF** | `python ai_tools_cli.py pdf-tools flatten -i <doc.pdf> [--strict]` |
| **Reorganize Pages with Blank Insertion** | `python ai_tools_cli.py pdf-tools reorganize -i <doc.pdf> -o <reorg.pdf> --plan "1,blank,2"` |
| **Set Excel Cell Value/Formula** | `python ai_tools_cli.py xlsx-set <file.xlsx> --sheet <sheet> --cell <cell> --value <val>` |
| **Insert Excel Rows & Shift Formulas** | `python ai_tools_cli.py xlsx-insert-rows <file.xlsx> --sheet <sheet> --start-row <row> --count <n>` |
| **Inspect Word Template Tokens** | `python ai_tools_cli.py docx-tokens <tpl.docx> [--json]` |
| **Atomic Word Text Replace** | `python ai_tools_cli.py docx-replace <doc.docx> --data <patch.json>` |
| **Excel Power Clean (VN Dates/CCCD/Dedup)** | `python ai_tools_cli.py xlsx-power-clean <file.xlsx> [--dedup MaNV] [--case title]` |
| **Excel Split / Merge Columns** | `python ai_tools_cli.py xlsx-transform-cols <file.xlsx> --mode split_vn_fullname --cols HoVaTen` |
| **Excel Pivot Summary & Chart** | `python ai_tools_cli.py xlsx-pivot-chart <file.xlsx> --group-by PhongBan --values Luong --agg SUM --chart bar` |
| **Excel No-Code XLOOKUP Join** | `python ai_tools_cli.py xlsx-xlookup-join <left.xlsx> --left-key MaNV --right-key MaNV --return-cols HoTen,Luong` |
| **Excel Split Workbook by Column (.zip)** | `python ai_tools_cli.py xlsx-split-col <file.xlsx> --col PhongBan -o split.zip` |
| **Excel Visual Cell Diff** | `python ai_tools_cli.py xlsx-diff <v1.xlsx> <v2.xlsx> -o diff.xlsx` |
| **Office Work & Daily Task Tracker** | `python ai_tools_cli.py xlsx-track {init,add,list,sync-daily,inspect}` |
| **Excel 57-Formula Playbook** | `python ai_tools_cli.py xlsx-formulas [--category lookup] [--export playbook.xlsx]` |
| **Word NĐ 30/2020 Polish & Page X/Y** | `python ai_tools_cli.py docx-polish <doc.docx> --preset nd30` |
| **Word Multi-File Merge** | `python ai_tools_cli.py docx-merge <doc1.docx> <doc2.docx> -o merged.docx` |
| **Word Split by Heading Level (.zip)** | `python ai_tools_cli.py docx-split <doc.docx> --level 1 -o chapters.zip` |
| **Word Watermark & Header/Footer Brand** | `python ai_tools_cli.py docx-watermark <doc.docx> --text "CONFIDENTIAL" --header "Company"` |
| **List System Hardware Printers** | `python ai_tools_cli.py print-list` |
| **Inspect Printer Status & Spooler Queue** | `python ai_tools_cli.py print-status [--printer <name>] [--json]` |
| **Enterprise Silent Print Document** | `python ai_tools_cli.py print-doc <file.pdf | docx> [--printer <name>] [--page-range "1-5, 8"] [--page-size a4] [--scaling fit] [--duplex vertical] [--n-up 2] [--grayscale]` |
| **Offline AST Excel Formula Evaluator** | `python ai_tools_cli.py xlsx-eval <file.xlsx> [-o <out.xlsx>] [--json]` |
| **Universal FTS5 RAG File Indexer** | `python ai_tools_cli.py rag-index <file> [--json]` |
| **Universal FTS5 RAG Semantic Search** | `python ai_tools_cli.py rag-search <query> [-k <top_k>] [-t <type>] [--json]` |
| **Hybrid AI Copilot Terminal Chat** | `python ai_tools_cli.py copilot-chat <prompt> [-p {hybrid,local,cloud}] [--index-file <file>] [--json]` |
| **Sync Core Docs Triad (MD->DOCX+PDF)** | `python tools/sync_core_docs.py [--build | --check]` |
| **Launch Desktop GUI** | `python main.py` |

---

## ⚡ Core Engine Architecture

```
antigravity-doc-handler/
├── core/                       # Layer 3: Core Document & Diagram Engines
│   ├── converter/              # Universal Converter, Decoupled MD, OCR, PDF Toolkit, COM Watchdog
│   ├── diagrams/               # Diagram Bridge, mxGraph, Code-to-Diagram, Editor, Legacy
│   ├── docx/                   # OpenXML Reader/Writer, Lists, Math, Diff, Mail Merge, Presets, Toolkit
│   ├── pptx/                   # 16:9 Widescreen Engine, DOCX-to-PPTX, Charts, Math
│   ├── xlsx/                   # Power Tools, Tracker, Universal Excel Engine, AST Evaluator, Validator, Streaming, 57-Formula Playbook, Reporting
│   ├── security/               # Cloudflare Document Auditor & Deep CDR / PII Sanitizer
│   ├── pipeline/               # Declarative Multi-Step Workflow Pipeline Engine
│   ├── copilot/                # Universal FTS5 RAG Vault & Hybrid AI Copilot
│   ├── printing/               # Enterprise Printing Engine (Vector /DeviceGray, Booklet Creep, DEVMODE Guard G2)
│   ├── env_doctor.py           # 12-Component Environment Health Doctor
│   └── system_maintenance.py   # 7-Pillar Enterprise System Maintenance & MPO Guard
├── server/                     # Layer 2: Web Studio REST API (:8000), Modular Routes, MCP Server (43 Tools), Persistent WAL Queue
├── cli/                        # Layer 2: Unified AI CLI Dispatcher & Subcommands
├── ui/                         # Layer 1: Desktop GUI (Tkinter) & Native WebView App Launcher
├── src/                        # Layer 1: Vite ES6 Web Studio Frontend & Modular Controllers
├── tools/                      # QA Diagnostics, Version Sync, Core Docs Sync & plantuml.jar
│   ├── unified_qa_diagnostic.py # 13-Gate Unified QA Diagnostic & Self-Repair Engine
│   ├── bump_version.py         # 10-Anchor Version Synchronization Tool (ERR_VER_001)
│   ├── sync_core_docs.py       # Core Docs Triad Compiler & Semantic Auditor (ERR_DOC_SYNC_001)
│   └── plantuml.jar            # Local PlantUML JAR (v1.2026+, bundled C4 standard library)
├── specs/                      # Declarative Diagram Specs (flow_*.json, master_erd_*.json)
├── tests/                      # Automated Test Suites (tests/test_*.py)
├── main.py                     # Official Entrypoint: Desktop GUI & Native Web App Mode
├── ai_tools_cli.py             # Official Entrypoint: Unified CLI Dispatcher for AI Agents
├── web_api_server.py           # Official Entrypoint: Web Studio REST API Server
├── mcp_server.py               # Official Entrypoint: Model Context Protocol STDIO Server
├── spec_diagram_engine.py      # Official Entrypoint: SVG Canvas Precision Diagram CLI
├── diagram_editor.py           # Official Entrypoint: Interactive Canvas Diagram Editor
└── .agents/                    # Antigravity Customizations (6 Skills & 13 Rules)
```

---

## 🎯 Important Invariants & Diagram Selection Rules

### Diagram Engine Selection Rules (MANDATORY)

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
