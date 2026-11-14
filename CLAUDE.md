# CLAUDE.md — Direct Intent-to-File Fast Dispatch

> **AI INSTRUCTION**: NEVER run recursive directory scans or blind grep. Jump DIRECTLY to the target file below.

## 🎯 Direct File Dispatch Matrix

| Intent / Feature Domain | Primary Source File | Primary Function / Symbol |
| --- | --- | --- |
| **System Diagnostics** | `core/env_doctor.py` | `run_doctor()` |
| **Universal Conversion** | `core/converter/engine.py` | `convert_universal()` |
| **PDF Toolkit & Security** | `core/converter/pdf_toolkit.py` | `PDFQuickToolkit`, `crypto_sign_pdf` |
| **Docx/PPTX AST Reader** | `core/docx/reader.py`, `core/pptx/reader.py` | `read_docx_to_json()`, `read_pptx()` |
| **Legal Track Changes Diff** | `core/docx/diff_engine.py` | `compare_docx_documents()` |
| **Mail Merge & Templating** | `core/docx/mail_merge.py`, `core/docx/writer.py` | `inspect_template_tokens()`, `patch_docx_template()` |
| **Unified 13-Gate QA Audit** | `core/qa/unified_audit.py` | `run_unified_qa()` |
| **Vietnamese OCR Engine** | `core/converter/ocr_pipeline.py` | `VietnameseOcrPipeline` |
| **Diagram Render (All)** | `core/diagrams/bridge.py` | `render_diagram_file_to_png()` |
| **Presentation Engine** | `core/pptx/docx_to_pptx.py`, `core/pptx/engine.py` | `convert_docx_to_pptx()` |
| **Excel In-Place Engine** | `core/xlsx/writer.py` | `UniversalExcelEngine`, `mutate_template_excel()` |
| **Excel Formula Evaluator** | `core/xlsx/formula_evaluator.py` | `ASTFormulaEvaluator.populate_cached_values()` |
| **Security Audit & Sanitizer**| `core/security/document_auditor.py` | `SecurityAuditor`, `OfficeSanitizer` |
| **Workflow Pipeline** | `core/pipeline/engine.py` | `WorkflowPipelineEngine` |
| **Code to Diagram** | `core/diagrams/code_to_diagram.py` | `generate_diagram_from_code()` |
| **Excel Power Suite** | `core/xlsx/power_tools.py` | `clean_and_standardize_workbook()`, `xlookup_join_workbooks()`, `build_pivot_summary_and_chart()` |
| **Office Work Tracker** | `core/xlsx/tracker.py` | `OfficeWorkTracker` |
| **Excel 57-Formula Playbook** | `core/xlsx/formula_playbook.py` | `FORMULA_CATALOG`, `generate_formula_playbook_xlsx()` |
| **Word Mastery Toolkit (NĐ30)** | `core/docx/toolkit.py` | `polish_docx_document()`, `merge_docx_files()`, `split_docx_by_heading()`, `apply_docx_watermark_and_hf()` |
| **VN Amount in Words** | `core/docx/design_system.py` | `number_to_vietnamese_words()`, `format_token_value()` |
| **Enterprise Silent Print** | `core/printing/printer_manager.py` | `PrinterManager.print_document()` |
| **Universal FTS5 RAG** | `core/copilot/universal_rag.py` | `UniversalRAGEngine` |
| **Hybrid AI Copilot** | `core/copilot/copilot_provider.py` | `CopilotOrchestrator` |
| **Periodic Maintenance** | `core/system_maintenance.py` | `run_periodic_maintenance()` |
| **CLI Dispatcher** | `ai_tools_cli.py`, `cli/` | `office_power_commands.py`, `v46_commands.py` |
| **MCP Server (43 Tools)** | `server/mcp_server.py` | `TOOL_DEFINITIONS`, `handle_rpc_request()` |
| **Web REST API (:8000)** | `server/app.py`, `server/routes/` | `xlsx_routes.py`, `docx_routes.py`, `doc_routes.py` |

## ⚡ Fast CLI Cheat Sheet

- Diagnostic Doctor: `python ai_tools_cli.py doctor`
- Quality QA Gate: `python ai_tools_cli.py qa-audit <file>`
- Excel Power Clean: `python ai_tools_cli.py xlsx-power-clean <file.xlsx> --dedup MaNV`
- Word NĐ 30/2020 Polish: `python ai_tools_cli.py docx-polish <doc.docx> --preset nd30`
- PDF Pro Suite (17 Tools): `python ai_tools_cli.py pdf-tools {merge,split,compress,watermark,giap-lai,crypto-sign,pdfa}`
- Run All Tests: `pytest -q`
- Version Audit: `python tools/bump_version.py --check`
