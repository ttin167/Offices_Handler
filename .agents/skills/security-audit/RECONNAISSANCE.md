# Phase 1: Reconnaissance Methodology

Reconnaissance establishes the ground truth of the target repository before any hunting begins. It produces two core artifacts: `architecture.md` and initial `coverage-ledger.json`.

## Objectives

1. Identify all **Trust Boundaries**:
   - Client-to-Server boundaries (HTTP endpoints, WebSocket, SSE).
   - Agent-to-Host boundaries (MCP tools, subprocess execution, CLI).
   - Untrusted Input-to-Parser boundaries (Office document uploads, user-provided JSON/YAML/XML/ZIP, image files).
2. Enumerate **Entry Points & Public APIs**:
   - REST API routes (e.g. `server/app.py`).
   - CLI subcommands (e.g. `ai_tools_cli.py`, `cli/ai_tools.py`).
   - MCP Tool declarations (e.g. `server/mcp_server.py`).
   - Static file handlers and upload sinks.
3. Discover **Third-Party & System Dependencies**:
   - Subprocess invocations (e.g. `soffice`, `plantuml.jar`, `chromium`, `tesseract`).
   - Native binary dependencies or C extensions (PyMuPDF `fitz`, Pillow).
4. Establish the **Deterministic Coverage Ledger**:
   - Break down the target repository into reviewable units (files or functional modules).
   - Initialize `coverage-ledger.json` tracking unit status: `unreviewed`, `in_progress`, `covered`.

## Output Format (`architecture.md`)

```markdown
# Target Architecture & Trust Boundary Map

## 1. System Overview
- Repository Root: <absolute_path>
- Primary Languages: Python, JavaScript/Node
- Runtime Framework: Standard library http.server, Vite, MCP STDIO

## 2. Trust Boundaries
| Boundary ID | Source (Untrusted) | Sink (Trusted / System) | Enforced Controls |
|---|---|---|---|
| TB-01 | Web Upload (`/api/upload`) | Local Disk (`.web_uploads/`) | Path sanitization, extension checks |
| TB-02 | Document Parse (`core/converter/`) | OpenXML / PyMuPDF Engine | ZipFile safe read, XML parsing |
| TB-03 | Subprocess (`tools/plantuml.jar`) | OS Process Execution | Arguments list escaping, timeout |

## 3. Review Ledger Units
- Total Files Identified: N
- Critical Units: server/app.py, core/security/, tools/unified_qa_diagnostic.py
```
