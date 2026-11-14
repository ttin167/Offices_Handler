# 📄 antigravity-doc-handler

**Universal Office Document Studio & Precision Technical Diagram Engine for Antigravity AI and Developers**

[![Python 3.10+](https://img.shields.io/badge/Python-3.10%2B-blue.svg)](https://www.python.org/)
[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)
[![Release: v5.0.0](https://img.shields.io/badge/Release-v5.0.0_Production-success.svg)](CHANGELOG.md)
[![Tests: 513 Passed](https://img.shields.io/badge/Tests-513%2F513_Passed-brightgreen.svg)](tests/)
[![MCP Tools: 43](https://img.shields.io/badge/MCP_Tools-43_Available-blue.svg)](server/mcp_server.py)
[![Antigravity Skills: 6](https://img.shields.io/badge/Antigravity_Skills-6_Specialized-indigo.svg)](.agents/skills/)
[![Antigravity IDE Ready](https://img.shields.io/badge/Antigravity_AI-Compatible-orange.svg)](AGENTS.md)
[![Resolution: 300+ DPI](https://img.shields.io/badge/DPI-300%2B_Vector_Grade-purple.svg)](spec_diagram_engine.py)

> 🚀 **Xem toàn bộ lộ trình tiến hóa & so sánh tính năng v1.0 → v5.0.0 tại [CHANGELOG.md](CHANGELOG.md)**.

---

## 🌟 Overview

`antigravity-doc-handler` is a specialized **Pure Office IT Automation Studio** designed for everyday office professionals, administrative specialists, finance/HR teams, technical writers, developers, and autonomous AI agents (especially within **Google Antigravity IDE**).

It brings together Word administrative formatting (Decree 30/2020/NĐ-CP & `vn_words`), no-code Excel data wrangling & dynamic formula evaluation, 17 Stirling-style PDF Pro tools, 16:9 PowerPoint deck synthesis, industrial vector printing, and publication-quality technical diagram generation.

### 🧠 Antigravity AI Skills & Customizations Suite

The workspace provides 6 specialized Antigravity Skills in `.agents/skills/` designed for autonomous AI agents and developers:

| Skill | Directory | Core Capabilities |
| --- | --- | --- |
| **doc-handler** | [`.agents/skills/doc-handler/`](.agents/skills/doc-handler/SKILL.md) | Bidirectional PDF ↔ Word ↔ Decoupled Markdown conversion, OpenXML table repair (`cantSplit`, `tblHeader`), and TOC dot-leader normalization. |
| **excel-automation** | [`.agents/skills/excel-automation/`](.agents/skills/excel-automation/SKILL.md) | Template-driven OpenXML Excel processing, format token extraction, live dynamic KPI linking, Excel Power Suite, Big Data streaming ($\le 50$MB RAM), and 4 edge-case guards (`ERR_XLSX_009`–`ERR_XLSX_014`). |
| **felo-slides** | [`.agents/skills/felo-slides/`](.agents/skills/felo-slides/SKILL.md) | 16:9 widescreen presentation synthesis, Slide Master inheritance, aspect-ratio locked diagram injection (`PPTX_INV_01`–`PPTX_INV_13`). |
| **mxgraph-diagram-engineering** | [`.agents/skills/mxgraph-diagram-engineering/`](.agents/skills/mxgraph-diagram-engineering/SKILL.md) | Draw.io (`.drawio` / native mxGraphModel XML) generation, orthogonal routing, monochrome academic line-art, fractional perimeter port anchoring (`MX_INV_01`–`MX_INV_20`). |
| **technical-diagrams** | [`.agents/skills/technical-diagrams/`](.agents/skills/technical-diagrams/SKILL.md) | Declarative 3-engine rendering (Canvas, Mermaid, PlantUML), 12 diagram types, 300+ DPI vector SVG / raster PNG, 1.6:1–1.85:1 aspect ratio, AABB collision avoidance, Word/DOCX auto-injection. |
| **security-audit** | [`.agents/skills/security-audit/`](.agents/skills/security-audit/SKILL.md) | Cloudflare-seeded 6-phase defensive security audit harness (Reconnaissance, Coverage-Led Hunting, Adversarial Validation, Structured Schema, Zero-Dependency Verification, Reporting) & document sanitizer (`ERR_SEC_001`–`007`). |

> 💡 **Auto-Sync Invariant (`ERR_DOC_SYNC_001` / `ERR_VER_001`)**: Khi có nâng cấp kỹ năng, thêm skill mới hoặc bump version, `README.md` sẽ được tự động đồng bộ huy hiệu (badges), bảng kỹ năng và các điểm nhấn phiên bản qua `tools/bump_version.py`.

### 🏢 3 Core Office IT Pillars (Tinh Hoa Tin Học Văn Phòng)

1. **Pillar 1 — Excel Power Suite, 57-Formula Playbook & Work Tracker (`core.xlsx`)**:
   - **6 Power Data Tools (`core.xlsx.power_tools`)**: VN Data Cleaner & Standardizer (trim spaces, normalize VN `dd/mm/yyyy` dates, force leading-zero `@` text on Phone/CCCD/Codes, deduplicate rows), Split Vietnamese Full Name (`Họ / Tên đệm / Tên`) & Merge Columns, Live `SUMIFS`/`COUNTIFS` Pivot Summary + Native Embedded Chart, No-Code XLOOKUP / INDEX-MATCH Join, Split Workbook by Column to `.zip`, and Visual Cell Diff (Amber/Emerald/Rose + `_Diff_Summary`).
   - **Office Work & Daily Task Tracker (`core.xlsx.tracker`)**: Full CRUD lifecycle & daily sync compatible with `project_tracking.xlsx` and `task_tracking.xlsx`.
   - **57-Formula Playbook & AST Evaluator**: Expanded 57 interactive Excel formulas across 6 categories + Pure-Python `ASTFormulaEvaluator` computing cached `<v>` tags offline.
2. **Pillar 2 — Word Mastery, Decree 30/2020/NĐ-CP & Vietnamese Amount in Words (`core.docx`)**:
   - **Vietnamese Financial Number Reader (`number_to_vietnamese_words`)**: Built-in Mail Merge filters (`{{SoTien:vn_words}}`, `{{SoTien:words_plain}}`, `{{NgayKy:date_vn}}`, `:vnd`, `:upper`, `:lower`, `:title`) plus header/footer table scanning.
   - **Word Mastery Toolkit (`core.docx.toolkit`)**: One-click Decree 30/2020/NĐ-CP or Executive document polishing (`polish_docx_document` with `ERR_DOCX_001/003/005` guards & Page X/Y), media-preserving multi-DOCX merge (`merge_docx_files`), Heading-level ZIP splitter (`split_docx_by_heading`), and native diagonal VML Watermark + Header/Footer branding (`apply_docx_watermark_and_hf`).
3. **Pillar 3 — Complete PDF Pro 17-Tool Suite (`core.converter.pdf_toolkit` & Web Studio)**:
   - Exposes all 17 Stirling-grade tools natively in the Web Studio UI and CLI/MCP: Merge, Split (with multi-file ZIP bundling), Compress, Diagonal Watermark, Images $\leftrightarrow$ PDF / 300 DPI ZIP Export, Interactive Visual Page Organizer, Bates / A4 Normalization, True Byte/Pixel Redaction (`ERR_PDF_001`), AES-256 Permission Protection, Multi-Page Overlap Stamp / Dấu Giáp Lai (`ERR_PDF_002`), Flatten & Scrub, Table Extraction to `@` Excel (`ERR_XLSX_007`), Digital Stamp Signing, AcroForms Inspect & Fill, PKCS#7 / PAdES Signing, and ISO 19005-2 PDF/A-2b archival conversion.

### 🆕 Highlights in v5.0.0 Production Grade (Enterprise Omni-Studio, AST Offline Formula Engine, Industrial Vector Print, In-Browser Viewers & Universal FTS5 RAG Copilot)

- **4-Phase Enterprise Omni-Studio Architecture & 4 Production Guards (v5.0.0)**:
  - *Phase 1 — AST Offline Formula Evaluator & COM Process Watchdog (`core.xlsx.formula_evaluator`, `core.converter.com_watchdog`, `server.routes`)*: Pure-Python AST evaluator computing 30+ Excel functions (`SUM`, `AVERAGE`, `COUNT`, `COUNTA`, `COUNTIF`, `COUNTIFS`, `SUMIF`, `SUMIFS`, `AVERAGEIF`, `AVERAGEIFS`, `IF`, `IFS`, `IFERROR`, `XLOOKUP`, `VLOOKUP`, `INDEX`, `MATCH`, `TEXTJOIN`, etc.) and injecting cached `<v>` tags directly into OpenXML (`populate_cached_values`); `ComProcessWatchdog` isolating Win32 COM threads with 20s timeout; modular `Router` singleton.
  - *Phase 2 — Industrial Vector Grayscale, Booklet Shingling Creep & Persistent WAL Queue (`core.printing`, `server.job_queue`)*: Native `/DeviceGray` PDF operator conversion preserving infinite vector text clarity; Saddle-Stitch Booklet Creep compensation ($\Delta x = k \times \text{thickness\_pt}$); **Guard G2 (`DeviceCapabilities DC_BINS`)** hardware tray capability check before `DEVMODE` mutation; **Guard G1 (`PRAGMA journal_mode = WAL; PRAGMA busy_timeout = 5000;`)** SQLite `PersistentJobQueue` with exponential backoff and crash recovery (`recover_stale_jobs`).
  - *Phase 3 — Inline Document Previews & Visual Field Mapper (`src/components/preview/InlinePreviewPane.tsx`, `src/features/word/MailMergePage.tsx`)*: Zero-CDN in-browser viewers for PDF, DOCX (AST A4 Paper Page), and XLSX (multi-sheet grid with `fx` formula bar) wrapped in **Guard G3 (Shimmer Skeleton Loading UX)**; interactive two-column Mail Merge token-to-column mapper with 3-tier Auto-Match and Live Record Preview.
  - *Phase 4 — Universal FTS5 RAG & Local-First AI Copilot Workspace (`core.copilot.universal_rag`, `core.copilot.copilot_provider`, `src/components/modals/CopilotSidebar.tsx`)*: Multi-format semantic chunking (PDF, DOCX, XLSX, CSV, JSON, MD) indexed in SQLite FTS5 (`unicode61 remove_diacritics 2`, $< 5\text{ms}$ latency); polymorphic `CopilotOrchestrator` locked by default to Local Ollama (`localhost:11434`) & Deterministic Offline RAG (with opt-in cloud gate `ANTIGRAVITY_ALLOW_CLOUD_AI=1`) plus interactive **Office Action Cards**.
- **6 Strategic Enterprise Subsystems & 5 Production Edge-Case Guards (v4.6.0)**:
  - *Subsystem 1 — Cryptographic PDF Signing (PKCS#7 / PAdES) & ISO 19005-2 PDF/A-2b (`core.converter.pdf_toolkit`)*: Embeds ISO `/Type /Sig` dictionaries with SHA-256 document digests, `.pfx`/`.p12` X.509 certificate support (`crypto_sign_pdf`, `verify_pdf_crypto_signature`), and converts documents into archival **PDF/A-2b** (`convert_to_pdfa`) with XMP `pdfaid:part=2`, `pdfaid:conformance=B`, and sRGB `OutputIntents`.
  - *Subsystem 2 — Two-Way Spreadsheet Reporting (`core.xlsx.reporting` & Guard #3 `ERR_XLSX_015`)*: Converts `.xlsx` workbooks directly into executive brief reports in `.docx`, `.md`, and `.pdf` (`excel_to_brief`) with Dual-Workbook formula evaluation (`data_only=True` + `data_only=False` fallback computing `SUM`, `AVERAGE`, `COUNT`, `MIN`, `MAX` when cached values are `None`) and multi-workbook consolidation (`consolidate_workbooks`).
  - *Subsystem 3 — Deep CDR & Smart Vietnamese PII Redactor (`core.security.sanitizer` & Guards #1 `ERR_SEC_006`, #2 `ERR_SEC_007`)*: Strips `vbaProject.bin` and PE (`MZ`) OLE payloads while automatically promoting `macroEnabled.main+xml` ContentTypes to standard OpenXML MIME types and cleaning `.rels`; redacts Vietnamese CCCD (validating province `001..096` + century/gender `0..3`), mobile numbers, context-anchored bank accounts (`STK`, `Số tài khoản`, `Account No`), MST, and emails without false-positive redaction on financial statement numbers (`150000000000 VND`).
  - *Subsystem 4 — Declarative Multi-Step Workflow Pipeline (`core.pipeline.engine` & Guard #4 `ERR_PIPE_001`)*: Chains multi-step document transformations via JSON/YAML specs (`WorkflowPipeline`, `run_bulk_directory_pipeline`) with memory-bounded concurrency (`MAX_JOB_WORKERS=2` default) and guaranteed `step_*_temp.*` cleanup in `try...finally`.
  - *Subsystem 5 — AST Code-to-Diagram & Native Draw.io SVG Exporter (`core.diagrams.code_to_diagram`)*: Parses Python AST and SQL/Prisma schemas into PlantUML, Mermaid, Draw.io `.drawio`, and pure-Python Vector `.svg` (`export_mxgraph_to_svg`) with automatic injection into both `.docx` and `.pptx`.
  - *Subsystem 6 — Unified Desktop Studio Bridge (`ui.webview_launcher` & Guard #5 `ERR_UI_001`)*: Launches the full Web Studio inside a native frameless Edge/Chrome `--app` window with an isolated `--user-data-dir` profile so `proc.wait()` tracks the exact window lifetime and cleanly shuts down the background server on exit.
- **Enterprise Excel Data Suite & Edge Cases Hardening (`core.xlsx`)**: Full compliance with invariants `ERR_XLSX_009` through `ERR_XLSX_015`:
  - *Data Loss & 15-Digit Limit Guard (`ERR_XLSX_009`)*: Enforces string `'@'` format for credit cards, CCCD (12 digits), and bank accounts (16–19 digits); regex parsing of Vietnamese dates (`26/05/2026`, `ngày 26 tháng 05 năm 2026`) into `datetime.date` + `DD/MM/YYYY`; converts accounting negatives `(100.000)` to `-100000` and dashes `-` to `0` with accounting format; normalizes Vietnamese locale decimals `20.000,50` $\rightarrow$ `20000.50`.
  - *XML Illegal Byte Sanitization & 32k Limit (`ERR_XLSX_010`)*: Strips ASCII control bytes `[\x00-\x08\x0b\x0c\x0e-\x1f]` preventing corrupt OpenXML workbooks; truncates strings $> 32,000$ characters with ellipsis.
  - *Modern Formula Compatibility (`ERR_XLSX_011`)*: Automatically prefixes dynamic array and recent formulas (`XLOOKUP`, `CONCAT`, `TEXTJOIN`, `UNIQUE`, `FILTER`, `SORT`, `SORTBY`, etc.) with `_xlfn.` preventing `#NAME?` in Excel 2016/2019; preserves Regex AST Reference Shifting (`ERR_XLSX_002`).
  - *Usability & Line Auto-Scaling (`ERR_XLSX_012`)*: Enforces visible gridlines `showGridLines = True`, AutoFilter range `auto_filter.ref`, and leaves data row heights as `None` so Excel dynamically fits multiline text without vertical clipping.
  - *Selective Formula & Cell Protection (`ERR_XLSX_013`)*: Automatically locks formula cells and headers (`locked=True`), leaving designated input ranges editable before enabling sheet protection.
  - *RAM-Constrained Big Data Streaming (`ERR_XLSX_014`)*: Module `core.xlsx.streaming` (`stream_read_xlsx` & `stream_write_xlsx`) maintains RAM consumption strictly $\le 50\text{MB}$ when processing workbooks exceeding 10 MB or 30,000 rows locally.
  - *6 Enterprise Subsystems*: Corporate pastel conditional formatting (red, green, yellow status highlights), in-cell Data Bars & 3-point Color Scales, Data Validation (dropdowns, CCCD 12-digit, year 2026), native in-sheet OpenXML charts (`BarChart`, `LineChart`, `PieChart` with safe `ERR_DIAG_003` placement), and Executive Index Dashboard with 2-way hyperlink navigation.
- **Excel Formula Playbook & Live Interactive Catalog (`core.xlsx.formula_playbook`)**: 57+ essential Excel functions documented across 4 comprehensive dimensions (Syntax, Sample formula, Real-world use case, Architectural reason why Excel chose this function); 1-click export to 2-sheet workbook (`Formula_Playbook` + `Live_Interactive_Formulas` with live calculating `=SUMIFS`, `=COUNTIFS`, `=AVERAGEIFS`, `=XLOOKUP`, `=IFS`, `=TEXTJOIN`).
- **Atomic Office Tools Suite**: Programmatic fine-grained inspection and manipulation of Word (.docx) and Excel (.xlsx) at cell/run level via MCP JSON-RPC 2.0 (43 tools) and CLI (`xlsx-set`, `xlsx-insert-rows`, `docx-tokens`, `docx-replace`, `xlsx-power-clean`, `docx-polish`), guaranteeing invariants `ERR_XLSX_004` (merged-cell safe targeting), `ERR_XLSX_007` (leading-zero string preservation `'@'`), and `ERR_DOCX_002` (run formatting preservation).
- **DOCX Quick Patch Studio**: Interactive Web Studio sub-tab with real-time **Two-Way Synchronization** between visual input form and JSON code editor, automatic AST token discovery, and one-click template patching.
- **PDF Pro Studio Suite (17 Tools)**: Complete Stirling-PDF alternative — Merge, Split ZIP, Compress, Watermark, Images $\leftrightarrow$ PDF, 300 DPI Export, Visual Organizer, true byte/pixel redaction (`ERR_PDF_001`), AES-256 permission encryption, $N$-cut overlapping page-spanning stamp (dấu giáp lai `ERR_PDF_002`), vector bake & raster flattening (`flatten`), Table Extract to Excel, rotation-compensated digital signing, AcroForms Inspect/Fill, PKCS#7/PAdES Crypto Sign, and ISO PDF/A-2b archival conversion.
- **Batch Mail Merge (`docx-mail-merge`)**: High-speed variable and table templating for Microsoft Word documents with nested block repetition and zero formatting loss.
- **Legal Redline Diff (`docx-diff`)**: Native OpenXML track changes comparison (`<w:ins>` / `<w:del>`) generating audit-compliant legal redlines between document revisions.
- **Vietnamese OCR & Scan Recovery (`ocr-scan`)**: Dual-pass HSV color filtering for red corporate seal extraction, deskew correction, and high-accuracy Tesseract OCR with `vie+eng` dictionaries.
- **Local-First Security & Web Studio API**: Hard-bound by default to `127.0.0.1:8000` with strict loopback CORS firewall, asynchronous SSE job queue (`server/job_manager.py`), REST endpoints for all tools, offline-locked AI Copilot default, and instant session purge (`/api/session/purge`).

### Why does this tool exist?

1. **Office Document Engines Lose Formatting**: Standard PDF-to-Word tools mangle Table of Contents (TOC) tab stops, split table rows awkwardly across page breaks, and fail to vertically center text.
2. **Diagram Engines Struggle with Aspect Ratios**: Traditional Mermaid Dagre layouts squeeze complex Hub & Spoke architectures into wide, unreadable horizontal strips (4:1 ratio) that turn into tiny unreadable specks when pasted into Word A4 pages.
3. **Decoupled Architecture**: Separate content (`.md`) from visual presentation (`.style.yaml`), allowing identical documents to be restyled effortlessly without modifying raw data.

---

## 🚀 Triple Core Diagram Engines & Universal Office Studio

### 1. 📐 Triple Core Diagram Architecture

`antigravity-doc-handler` unites **3 rendering engines** behind a declarative **JSON Single Source of Truth**:

- **Canvas Engine (`spec_diagram_engine.py`)**: Pixel-perfect Cartesian coordinates, orthogonal Manhattan routing, and white-halo collision avoidance for complex screen flows.
- **PlantUML Engine (`plantuml_renderer.py`)**: Industry-standard UML for Use Case, C4 Architecture, Component/Deployment, Activity Swimlane, Class OOP, and Package diagrams using bundled `plantuml.jar` and Java runtime.
- **Mermaid Engine (`mermaid_renderer.py`)**: Fast dynamic text-based rendering for ERD (Database Schema 3–15 tables), Sequence diagrams, Flowcharts, State machines, and Mind Maps.

#### 📊 12 Supported Diagram Types & Engine Mapping

| # | Loại Diagram | Engine đã chốt | Vai trò chính | Lý do thực tế |
| --- | --- | --- | --- | --- |
| 1 | **Use Case Diagram** | `plantuml` | Architect / Analyst | Actor người que, quan hệ `<<include>>`, `<<extend>>` chuẩn UML |
| 2 | **C4 Architecture** | `plantuml` | Architect | Thư viện C4 chuẩn quốc tế (`<C4/C4_Context>`), không vỡ dây container |
| 3 | **Component / Deployment** | `plantuml` | Dev / DevOps | Đúng hình khối 3D `node`, `database`, `component` chuẩn kiến trúc |
| 4 | **Activity Swimlane** | `plantuml` | Operations / Analyst | Phân làn cột (`\|Partition\|`) thẳng đứng tuyệt đối, không đè dây |
| 5 | **Class Diagram** | `plantuml` | Dev | Hỗ trợ trọn vẹn OOP (`+`, `-`, `#`, generics `<T>`, composition) |
| 6 | **Package Diagram** | `plantuml` | Dev / Architect | Hỗ trợ stereotype `<<Folder>>` trực quan cho cấu trúc thư mục |
| 7 | **ERD (Database Schema)** | `mermaid` | Dev / Data | Nền phẳng pastel hiện đại, gọn gàng cho phân hệ 3–15 bảng |
| 8 | **Sequence Diagram** | `mermaid` | Dev / Architect | Chuỗi gọi API thanh thoát, đánh số tự động `autonumber` |
| 9 | **Flowchart / Process** | `mermaid` | Operations / Office | Bẻ nhánh if/else tự do, đổi màu khối nhanh bằng CSS/style |
| 10 | **State Diagram** | `mermaid` | Dev | Trạng thái bo góc tròn, màu sắc hiện đại hơn nét vẽ thô |
| 11 | **Mind Map** | `mermaid` | Office / Dev | Phân rã đầu việc nhanh, màu pastel chia nhánh trực quan |
| 12 | **Screen Flow (Interactive)** | `canvas` | Dev / Designer | Định vị X,Y pixel-perfect, kéo thả trên `diagram_editor.py` |

### 2. 🛡️ Windows Production Invariants for PlantUML

1. **C4 Standard Library**: Uses internal `<C4/C4_Context>` (100% offline, never depends on raw github URLs).
2. **UTF-8 Subprocess**: Explicit `-charset UTF-8` JVM argument preventing Vietnamese text corruption.
3. **16K Pixel Limit**: `-DPLANTUML_LIMIT_SIZE=16384` eliminates truncating on massive high-DPI ERDs.
4. **Auto Word Injection**: Automatically injects diagrams into target Word `.docx` documents.

### 3. 🎨 Interactive Canvas Diagram Editor (`diagram_editor.py`)

- **Direct Visual Editing**: Drag & drop screen nodes, modify dimensions, adjust action labels, and connect ports visually.
- **Two-Way JSON Synchronization**: JSON spec is the Single Source of Truth. Changes are saved back to clean JSON with zero data loss.
- **Anti-Flicker Drag & Drop (Tkinter Tag Move)**: Uses `canvas.move` and recalculates only incident edges during motion, eliminating canvas flashes and lag.
- **Safe Node ID Renaming (Collision Guard)**: Strict duplicate ID validation before updating edge references, preventing corrupted state.
- **Space-Pan Keyboard Isolation**: Canvas panning via Space key strictly verifies canvas focus, allowing seamless typing in property fields.
- **Precise Edge Hit-Testing**: Uses `canvas.find_overlapping` bounding box to effortlessly select and edit orthogonal connection lines.
- **30-Step Undo / Redo**: Deep state snapshot history (`Ctrl+Z` / `Ctrl+Y`).
- **High-Resolution PNG Export**: One-click 300+ DPI export (`scale=3`) with automatic system preview.
- **Embedded in GUI & Standalone CLI**: Accessible directly from `main.py` or standalone via `python diagram_editor.py --spec <file.json>`.

### 4. 📑 Universal Office Document Studio (`smart_post_processor.py`, `converter_engine.py`)

- **100% Fidelity PDF ↔ DOCX**: Converts PDF to editable Word while applying **Smart Post-Processor v6** to enforce strict OpenXML table invariants.
- **Word Table OpenXML Invariants**:
  - `<w:cantSplit/>`: Prohibits table rows from splitting across page breaks.
  - `<w:tblHeader/>`: Automatically repeats header rows across multiple pages.
  - `<w:vAlign w:val="center"/>`: Vertically centers text inside table cells.
  - `<w:shd w:fill="FFE8E0"/>`: Applies elegant Peach shading to table header rows.
- **TOC Dot-Leader Normalization**: Native right-aligned tab stops with leader dots (`.......`).
- **Decoupled Markdown + Style YAML**: Exports clean Markdown without inline CSS or YAML frontmatter clutter, pairing it with a standalone `.style.yaml` stylesheet.

### 5. 🌐 Universal Web Studio (`npm run dev`)

Antigravity Office Studio features a full-fledged, privacy-first local web application runnable directly on your machine (`127.0.0.1:8000` / `:5173`):

- **Presentation Studio**: Automatic Widescreen 16:9 slide synthesis from Word (.docx) and PDF (.pdf) documents with 4 themes (`thesis_blue`, `corporate_blue`, `modern_dark`, `academic_light`), live card inspector, and PPTX/Spec download.
- **Universal Document Hub**: Multi-directional conversion matrix (PDF ↔ Word ↔ Markdown ↔ PPTX) with live progress tracking.
- **Diagram Studio**: Interactive live renderer for Mermaid, PlantUML, and Draw.io mxGraph models.
- **Single-Command Launch**: Run `npm run dev` to simultaneously launch the Python REST API Bridge (port 8000) and Vite Web App (port 5173).

---

## 📦 Installation

### Prerequisites

- Python 3.10 or higher
- Node.js >= 18 (for Web Studio and Mermaid rendering)
- Java JRE/JDK >= 8 (for PlantUML diagrams)
- Microsoft Edge or Google Chrome (for headless Canvas diagram rasterization)

### Setup

```bash
# Clone repository
git clone https://github.com/YukiiMin/antigravity-doc-handler.git
cd antigravity-doc-handler

# Install Python dependencies
pip install -r requirements.txt

# Install Web Studio frontend dependencies
npm install
```

### Launch Web Studio (Localhost)

```bash
# Launches both Python API Bridge (:8000) and Vite UI (:5173)
npm run dev

# Or build the production web bundle:
npm run build
```

---

## 💻 CLI Usage Guide

`antigravity-doc-handler` provides a unified command line interface via `ai_tools_cli.py`:

### 1. Unified Multi-Engine Diagram Rendering (Recommended)

Auto-detects backend (`canvas`, `mermaid`, `plantuml`) from the JSON spec's `"engine"` field:

```bash
python ai_tools_cli.py diagram-render spec.json -o diagram.png
```

### 2. Dedicated Diagram Renderers

```bash
# Render PlantUML (ERD, Class, C4, Swimlane, Use Case, Mind Map)
python ai_tools_cli.py plantuml-render erd_spec.json -o erd.png --dpi 300

# Render Mermaid (Sequence, Flowchart, State)
python ai_tools_cli.py mermaid-render sequence_spec.json -o seq.png

# Render SVG Canvas (Screen Flow)
python ai_tools_cli.py spec-render screen_flow_spec.json -o screen_flow.png --scale 3
```

### 3. Word & PDF to PowerPoint Presentation (16:9 Widescreen)

```bash
# Convert Word DOCX to 16:9 Presentation PPTX
python ai_tools_cli.py docx-to-pptx document.docx -o presentation.pptx --theme thesis_blue

# Convert PDF to 16:9 Presentation PPTX (with optional page range)
python ai_tools_cli.py pdf-to-pptx document.pdf -o presentation.pptx --theme modern_dark --pages 1-15
```

### 4. Document Conversion Matrix

```bash
# Convert PDF to DOCX (with Smart Post-Processor v6)
python -m ai_tools_cli convert input.pdf -o output.docx

# Convert DOCX to Decoupled Markdown + Style YAML
python -m ai_tools_cli convert document.docx -t md

# Convert Markdown to DOCX using style tokens
python -m ai_tools_cli convert document.md -t docx --style document.style.yaml

# Convert DOCX to PDF (Word COM Automation)
python -m ai_tools_cli convert document.docx -t pdf
```

### 4. Document Inspection & Diagram Insertion

```bash
# Inspect Word document headings, tables, and OpenXML properties
python -m ai_tools_cli inspect-doc document.docx

# Insert diagram image directly beneath a specific heading
python -m ai_tools_cli insert-diagram document.docx screen_flow.png \
  -s "3.1.2 Android User Screen Flow" \
  -c "Figure 3.1: Android User Screen Flow" \
  -w 14.0
```

### 5. Interactive Canvas Diagram Editor

Launch the visual two-way editor for real-time drag-and-drop node manipulation, connection routing, and 300+ DPI PNG export:

```bash
# Open editor with an existing diagram spec:
python diagram_editor.py --spec specs/flow_android_user_v2_spec.json

# Or open a blank canvas:
python diagram_editor.py
```

### 6. Desktop Drag-and-Drop GUI

For non-technical users, launch the native desktop studio with integrated **[📐 Canvas Diagram Editor]** button:

```bash
python main.py
# Or double-click run_gui.bat on Windows
```

### 7. Atomic Office Tools (Word & Excel Manipulation)

Programmatic cell and token editing for terminal users and automated AI pipelines:

```bash
# Update Excel cell with merged-cell safety and formula support (ERR_XLSX_004, ERR_XLSX_007)
python ai_tools_cli.py xlsx-set data.xlsx --sheet "Report" --cell "C7" --value "100" --formula "=SUM(C2:C6)"

# Expand table rows with prototype row style cloning and formula shifting (ERR_XLSX_001, ERR_XLSX_002)
python ai_tools_cli.py xlsx-insert-rows data.xlsx --sheet "Report" --start-row 10 --count 5 --proto 9

# Inspect Word template placeholders and variables (JSON AST)
python ai_tools_cli.py docx-tokens template.docx --json

# In-place Word text replacement preserving run formatting (ERR_DOCX_002)
python ai_tools_cli.py docx-replace contract.docx --data '{"TEN_KHACH_HANG": "Nguyen Van A"}'
```

---

## 🤖 Antigravity AI Agent Integration

This repository is built natively for AI Agents operating in **Google Antigravity IDE**, Cursor, or Claude Code.

### Included Customizations

- `.agents/skills/doc-handler/SKILL.md`: Instructs AI agents on converting documents, fixing table splits, and formatting headings.
- `.agents/skills/technical-diagrams/SKILL.md`: Instructs AI agents on composing declarative JSON specs, calculating zero-overlap layouts, and routing lines.
- `.agents/rules/rule_decoupled_document_converter.md`: System invariants for document fidelity and diagram styling.
- `AGENTS.md`: Full AI agent operational guide.

### Sample AI Prompt
>
> *"Extract the Android User Screen Flow diagram from page 322 of Report3.docx, reverse-engineer its topology into specs/flow_android_user_v2_spec.json, render a 300+ DPI high-resolution PNG, and re-inject it back under Section 3.1.2 with proper APA captioning."*

---

## 📐 Declarative Diagram Spec Format

Below is a minimal sample of the JSON schema used by `spec_diagram_engine.py`:

```json
{
  "width": 1320,
  "height": 720,
  "font_family": "Segoe UI, -apple-system, Roboto, sans-serif",
  "font_size": 10.5,
  "bg_color": "#ffffff",
  "scale": 3,
  "nodes": [
    {
      "id": "home",
      "label": "Member Home Screen",
      "x": 360,
      "y": 265,
      "width": 150,
      "height": 44,
      "type": "primary"
    },
    {
      "id": "profile",
      "label": "Profile Screen",
      "x": 550,
      "y": 460,
      "width": 115,
      "height": 42,
      "type": "standard"
    }
  ],
  "edges": [
    {
      "source": "home",
      "target": "profile",
      "source_port": "bottom",
      "target_port": "left",
      "source_offset": 30,
      "label": "Bottom Nav",
      "line_style": "solid",
      "waypoints": [[520, 309], [520, 481]],
      "label_pos": 0.45,
      "label_offset_x": -16
    }
  ]
}
```

---

## 📁 Project Structure

```
antigravity-doc-handler/
├── .agents/                                     # Antigravity Agent Configuration
│   ├── rules/                                   # 13 Universal Engineering Rules
│   └── skills/                                  # 6 Pure Office Agent Workflow Skills
├── core/                                        # Layer 3: Core Document, Spreadsheet & Diagram Engines
│   ├── converter/                               # Universal Converter, Markdown, OCR, PDF Toolkit (17 Actions)
│   ├── copilot/                                 # Hybrid AI Copilot & Universal FTS5 RAG Vault (Pure Office IT)
│   ├── diagrams/                                # Diagram Bridge, mxGraph, Code-to-Diagram, Editor, Legacy
│   ├── docx/                                    # OpenXML Reader/Writer, Mail Merge, VN Words, NĐ30 Toolkit
│   ├── pptx/                                    # 16:9 Widescreen Engine, DOCX-to-PPTX, Charts, Math
│   ├── xlsx/                                    # Universal Excel Engine, Power Tools, 57 Formulas, Tracker
│   ├── security/                                # Cloudflare Document Auditor & Deep CDR / PII Sanitizer
│   ├── pipeline/                                # Declarative Multi-Step Workflow Pipeline Engine
│   ├── printing/                                # Enterprise Silent Printing Engine
│   ├── qa/                                      # Unified QA Audit Wrapper
│   └── env_doctor.py                            # 12-Component Environment Health Doctor
├── server/                                      # Layer 2: Protocol & Service Layer
│   ├── app.py                                   # Web Studio REST API Server (:8000) & Security Guards
│   ├── routes/                                  # Modular Routers (xlsx_routes.py, docx_routes.py)
│   ├── mcp_server.py                            # Model Context Protocol Server (43 Tools, JSON-RPC 2.0)
│   ├── job_manager.py                           # Asynchronous SSE Background Task Queue
│   └── db.py                                    # Zero-Config SQLite History Storage
├── cli/                                         # Layer 2: Command-Line Interface Layer
│   ├── ai_tools.py                              # Unified CLI Dispatcher
│   ├── office_power_commands.py                 # Pure Office Power CLI Subcommands (XLSX/DOCX/Tracker/Copilot)
│   └── converter_cli.py                         # Legacy Converter CLI Handler
├── ui/                                          # Layer 1: Desktop GUI & Native WebView Bridge
│   ├── desktop_app.py                           # Tkinter Native Desktop Studio
│   └── webview_launcher.py                      # Isolated Edge/Chrome App-Mode Window Launcher
├── src/                                         # Layer 1: React 18 + TypeScript + Tailwind v3 Studio
│   ├── App.tsx / main.tsx / styles/             # Shell Router, Command Palette (Ctrl+K) & Theme Tokens
│   ├── context/ & services/                     # Workspace/Session/Toast State & Typed Local REST Client
│   ├── components/                              # SplitWorkspace, InlinePreviewPane & Global Modals
│   └── features/                                # Pure Office Feature Workspaces (home, word, pdf, excel, convert, visual)
├── tools/                                       # Diagnostics, Version Sync, Tracker & Diagram Generators
│   ├── unified_qa_diagnostic.py                 # 13-Gate Unified QA Diagnostic & Self-Repair Engine
│   ├── bump_version.py                          # 10-Anchor Version Synchronization Tool (ERR_VER_001)
│   ├── sync_core_docs.py                        # Core Documentation Triad Compiler & Semantic Auditor (ERR_DOC_SYNC_001)
│   ├── project_tracker.py                       # Office Work Tracker & Sync CLI (powered by core.xlsx.tracker)
│   ├── format_diff_excel.py                     # 12-Gate Excel Format Comparison Tool
│   ├── gen_*.py                                 # Reference Diagram Generators (ERD, Hub & Spoke, Use Case)
│   └── plantuml.jar                             # Local PlantUML Runtime (Bundled C4 Standard Library)
├── docs/                                        # Core Technical & User Documentation Triad (.md, .docx, .pdf)
├── specs/                                       # Declarative Diagram Specifications
├── presets/                                     # Decoupled Document Style Presets (.style.yaml)
├── tests/                                       # Automated Unit & Integration Test Suites
├── main.py                                      # Official Entrypoint: Desktop GUI & Native Web App
├── ai_tools_cli.py                              # Official Entrypoint: Unified AI & Terminal CLI
├── web_api_server.py                            # Official Entrypoint: Web Studio REST API Server
├── mcp_server.py                                # Official Entrypoint: MCP STDIO Server
├── spec_diagram_engine.py                       # Official Entrypoint: SVG Canvas Diagram CLI
├── diagram_editor.py                            # Official Entrypoint: Interactive Canvas Editor
├── requirements.txt / pyproject.toml / package.json # Dependency & Build Manifests
├── ARCHITECTURE.md / CHANGELOG.md / AGENTS.md   # Architectural & Agent Operational Manuals
└── README.md                                    # Repository Overview
```

---

## 📄 License

This project is licensed under the [MIT License](LICENSE).
Feel free to use, modify, distribute, and integrate into your own workflows.
