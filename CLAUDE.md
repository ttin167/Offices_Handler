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
| **BA Story / Gherkin Linter**| `core/ba_hub/story_builder.py` | `UserStoryBuilder`, `lint_story()` |
| **RTM Traceability Matrix** | `core/ba_hub/rtm_matrix.py` | `RTMMatrixEngine` |
| **DBML / SQL Schema 2-Way** | `core/ba_hub/dbml_converter.py` | `DBMLConverter` |
| **Satsuma Lineage Graph** | `core/ba_hub/satsuma_engine.py` | `SatsumaMappingEngine` |
| **7 SQL Integrity Tests** | `core/ba_hub/sql_integrity.py` | `SQLIntegritySuite` |
| **Obsidian Vault Compiler** | `core/ba_hub/obsidian_vault.py` | `ObsidianVaultExporter`, `ObsidianVaultCompiler` |
| **OpenAPI Data Dictionary** | `core/ba_hub/data_dictionary.py` | `DataDictionaryScanner` |
| **Decision Table Matrix** | `core/ba_hub/decision_table.py` | `DecisionTableMatrix` |
| **Meeting Notes to Stories** | `core/ba_hub/meeting_parser.py` | `MeetingNotesParser` |
| **PlantUML Salt Wireframe** | `core/diagrams/salt_wireframe.py` | `SaltWireframeEngine` |
| **Labor Law Search & FTS5** | `core/ba_hub/labor_law_search.py` | `LaborLawSearchEngine` |
| **Labor Compliance Linter** | `core/ba_hub/labor_law_compliance.py`| `LaborLawComplianceLinter` |
| **Regulatory Law Watcher** | `core/ba_hub/regulatory_watcher.py` | `RegulatoryWatcher` |
| **Enterprise Silent Print** | `core/printing/printer_manager.py` | `PrinterManager.print_document()` |
| **Universal FTS5 RAG** | `core/ba_hub/universal_rag.py` | `UniversalRAGVault` |
| **Hybrid AI Copilot** | `core/ba_hub/copilot_provider.py` | `CopilotHybridOrchestrator` |
| **HR Governance & Scopes** | `core/ba_hub/hr_governance.py` | `HRDataScope`, `SoDConflictAuditor` |
| **BABOK URD Synthesizer** | `core/ba_hub/urd_generator.py` | `URDGenerator`, `URDLinter` |
| **Functional Spec Doc (FSD)** | `core/ba_hub/fsd_generator.py` | `FSDGenerator`, `FSDLinter` |
| **Periodic Maintenance** | `core/system_maintenance.py` | `run_periodic_maintenance()` |
| **CLI Dispatcher** | `ai_tools_cli.py`, `cli/` | Entrypoint & Subcommands |
| **MCP Server (50 Tools)** | `server/mcp_server.py` | `TOOL_DEFINITIONS`, `handle_rpc_request()` |
| **Web REST API (:8000)** | `server/app.py`, `server/routes/` | Modular Flask/FastAPI Endpoints |

## ⚡ Fast CLI Cheat Sheet

- Diagnostic Doctor: `python ai_tools_cli.py doctor`
- Quality QA Gate: `python ai_tools_cli.py qa-audit <file>`
- HR Governance: `python ai_tools_cli.py hr-governance data-scopes`
- URD Generation: `python ai_tools_cli.py ba-urd-generate <spec.json>`
- FSD Generation: `python ai_tools_cli.py ba-fsd-generate <spec.json>`
- Run All Tests: `pytest -q` (512 tests, 100% pass rate)
- Version Audit: `python tools/bump_version.py --check`
