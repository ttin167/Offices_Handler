# 📄 antigravity-doc-handler

**Universal Office Document Studio & Precision Technical Diagram Engine for Antigravity AI and Developers**

[![Python 3.10+](https://img.shields.io/badge/Python-3.10%2B-blue.svg)](https://www.python.org/)
[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)
[![Release: v4.3.0](https://img.shields.io/badge/Release-v4.3.0_Production-success.svg)](CHANGELOG.md)
[![Tests: 229 Passed](https://img.shields.io/badge/Tests-229%2F229_Passed-brightgreen.svg)](tests/)
[![MCP Tools: 20](https://img.shields.io/badge/MCP_Tools-20_Available-blue.svg)](server/mcp_server.py)
[![Antigravity Skills: 5](https://img.shields.io/badge/Antigravity_Skills-5_Specialized-indigo.svg)](.agents/skills/)
[![Antigravity IDE Ready](https://img.shields.io/badge/Antigravity_AI-Compatible-orange.svg)](AGENTS.md)
[![Resolution: 300+ DPI](https://img.shields.io/badge/DPI-300%2B_Vector_Grade-purple.svg)](spec_diagram_engine.py)

> 🚀 **Xem toàn bộ lộ trình tiến hóa & so sánh tính năng v1.0 → v4.3.0 tại [CHANGELOG.md](CHANGELOG.md)**.

---

## 🌟 Overview

`antigravity-doc-handler` is a specialized Python toolkit designed for software engineers, business analysts, technical writers, and autonomous AI agents (especially within **Google Antigravity IDE**). 

It bridges the gap between raw document conversion, OpenXML Word standard compliance, and publication-quality technical diagram generation.

### 🧠 Antigravity AI Skills & Customizations Suite

The workspace provides 5 specialized Antigravity Skills in `.agents/skills/` designed for autonomous AI agents and developers:

| Skill | Directory | Core Capabilities |
|---|---|---|
| **doc-handler** | [`.agents/skills/doc-handler/`](.agents/skills/doc-handler/SKILL.md) | Bidirectional PDF ↔ Word ↔ Decoupled Markdown conversion, OpenXML table repair (`cantSplit`, `tblHeader`), and TOC dot-leader normalization. |
| **excel-automation** | [`.agents/skills/excel-automation/`](.agents/skills/excel-automation/SKILL.md) | Template-driven OpenXML Excel processing, format token extraction, live dynamic KPI linking, 6 enterprise subsystems, Big Data streaming ($\le 50$MB RAM), and 4 edge-case guards (`ERR_XLSX_009`–`ERR_XLSX_014`). |
| **felo-slides** | [`.agents/skills/felo-slides/`](.agents/skills/felo-slides/SKILL.md) | 16:9 widescreen presentation synthesis, Slide Master inheritance, aspect-ratio locked diagram injection (`PPTX_INV_01`–`PPTX_INV_13`). |
| **mxgraph-diagram-engineering** | [`.agents/skills/mxgraph-diagram-engineering/`](.agents/skills/mxgraph-diagram-engineering/SKILL.md) | Draw.io (`.drawio` / native mxGraphModel XML) generation, orthogonal routing, monochrome academic line-art, fractional perimeter port anchoring (`MX_INV_01`–`MX_INV_20`). |
| **technical-diagrams** | [`.agents/skills/technical-diagrams/`](.agents/skills/technical-diagrams/SKILL.md) | Declarative 3-engine rendering (Canvas, Mermaid, PlantUML), 12 diagram types, 300+ DPI vector SVG / raster PNG, 1.6:1–1.85:1 aspect ratio, AABB collision avoidance, Word/DOCX auto-injection. |

> 💡 **Auto-Sync Invariant (`ERR_DOC_SYNC_001` / `ERR_VER_001`)**: Khi có nâng cấp kỹ năng, thêm skill mới hoặc bump version, `README.md` sẽ được tự động đồng bộ huy hiệu (badges), bảng kỹ năng và các điểm nhấn phiên bản qua `tools/bump_version.py`.

### 🆕 Highlights in v4.3.0 Production Grade (Formula Playbook & Enterprise Excel Suite)
- **Enterprise Excel Data Suite & Edge Cases Hardening (`core.xlsx`)**: Full compliance with invariants `ERR_XLSX_009` through `ERR_XLSX_014`:
  - *Data Loss & 15-Digit Limit Guard (`ERR_XLSX_009`)*: Enforces string `'@'` format for credit cards, CCCD (12 digits), and bank accounts (16–19 digits); regex parsing of Vietnamese dates (`26/05/2026`, `ngày 26 tháng 05 năm 2026`) into `datetime.date` + `DD/MM/YYYY`; converts accounting negatives `(100.000)` to `-100000` and dashes `-` to `0` with accounting format; normalizes Vietnamese locale decimals `20.000,50` $\rightarrow$ `20000.50`.
  - *XML Illegal Byte Sanitization & 32k Limit (`ERR_XLSX_010`)*: Strips ASCII control bytes `[\x00-\x08\x0b\x0c\x0e-\x1f]` preventing corrupt OpenXML workbooks; truncates strings $> 32,000$ characters with ellipsis.
  - *Modern Formula Compatibility (`ERR_XLSX_011`)*: Automatically prefixes dynamic array and recent formulas (`XLOOKUP`, `CONCAT`, `TEXTJOIN`, `UNIQUE`, `FILTER`, `SORT`, `SORTBY`, etc.) with `_xlfn.` preventing `#NAME?` in Excel 2016/2019; preserves Regex AST Reference Shifting (`ERR_XLSX_002`).
  - *Usability & Line Auto-Scaling (`ERR_XLSX_012`)*: Enforces visible gridlines `showGridLines = True`, AutoFilter range `auto_filter.ref`, and leaves data row heights as `None` so Excel dynamically fits multiline text without vertical clipping.
  - *Selective Formula & Cell Protection (`ERR_XLSX_013`)*: Automatically locks formula cells and headers (`locked=True`), leaving designated input ranges editable before enabling sheet protection.
  - *RAM-Constrained Big Data Streaming (`ERR_XLSX_014`)*: Module `core.xlsx.streaming` (`stream_read_xlsx` & `stream_write_xlsx`) maintains RAM consumption strictly $\le 50\text{MB}$ when processing workbooks exceeding 10 MB or 30,000 rows on Render.com 512MB RAM free tier.
  - *6 Enterprise Subsystems*: Corporate pastel conditional formatting (red, green, yellow status highlights), in-cell Data Bars & 3-point Color Scales, Data Validation (dropdowns, CCCD 12-digit, year 2026), native in-sheet OpenXML charts (`BarChart`, `LineChart`, `PieChart` with safe `ERR_DIAG_003` placement), and Executive Index Dashboard with 2-way hyperlink navigation.
- **Excel Formula Playbook & Live Interactive Catalog (`core.xlsx.formula_playbook`)**: 32+ essential Excel functions documented across 4 comprehensive dimensions (Syntax, Sample formula, Real-world use case, Architectural reason why Excel chose this function); 1-click export to 2-sheet workbook (`Formula_Playbook` + `Live_Interactive_Formulas` with live calculating `=SUMIFS`, `=COUNTIFS`, `=AVERAGEIFS`, `=XLOOKUP`, `=IFS`, `=TEXTJOIN`).
- **Atomic Office Tools Suite (8 AI Agent Tools)**: Programmatic fine-grained inspection and manipulation of Word (.docx) and Excel (.xlsx) at cell/run level via MCP JSON-RPC 2.0 (20 tools) and CLI (`xlsx-set`, `xlsx-insert-rows`, `docx-tokens`, `docx-replace`), guaranteeing invariants `ERR_XLSX_004` (merged-cell safe targeting), `ERR_XLSX_007` (leading-zero string preservation `'@'`), and `ERR_DOCX_002` (run formatting preservation).
- **DOCX Quick Patch Studio**: Interactive Web Studio sub-tab with real-time **Two-Way Synchronization** between visual input form and JSON code editor, automatic AST token discovery, and one-click template patching.
- **PDF Pro Studio Suite (16 Tools)**: Complete Stirling-PDF alternative — true byte/pixel redaction (`ERR_PDF_001`), AES-256 permission encryption, $N$-cut overlapping page-spanning stamp (dấu giáp lai `ERR_PDF_002`), vector bake & raster flattening (`flatten`), visual page organizer with drag & drop, 300 DPI image export, and rotation-compensated digital signing.
- **Batch Mail Merge (`docx-mail-merge`)**: High-speed variable and table templating for Microsoft Word documents with nested block repetition and zero formatting loss.
- **Legal Redline Diff (`docx-diff`)**: Native OpenXML track changes comparison (`<w:ins>` / `<w:del>`) generating audit-compliant legal redlines between document revisions.
- **Vietnamese OCR & Scan Recovery (`ocr-scan`)**: Dual-pass HSV color filtering for red corporate seal extraction, deskew correction, and high-accuracy Tesseract OCR with `vie+eng` dictionaries.
- **Render Cloud Ready & Web Studio API**: Asynchronous SSE job queue (`server/job_manager.py`), REST endpoints for all tools, and live multi-format preview.


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
|---|---|---|---|---|
| 1 | **Use Case Diagram** | `plantuml` | BA | Actor người que, quan hệ `<<include>>`, `<<extend>>` chuẩn UML |
| 2 | **C4 Architecture** | `plantuml` | Architect | Thư viện C4 chuẩn quốc tế (`<C4/C4_Context>`), không vỡ dây container |
| 3 | **Component / Deployment** | `plantuml` | Dev / DevOps | Đúng hình khối 3D `node`, `database`, `component` chuẩn kiến trúc |
| 4 | **Activity Swimlane** | `plantuml` | BA | Phân làn cột (`\|Partition\|`) thẳng đứng tuyệt đối, không đè dây |
| 5 | **Class Diagram** | `plantuml` | Dev | Hỗ trợ trọn vẹn OOP (`+`, `-`, `#`, generics `<T>`, composition) |
| 6 | **Package Diagram** | `plantuml` | Dev / Architect | Hỗ trợ stereotype `<<Folder>>` trực quan cho cấu trúc thư mục |
| 7 | **ERD (Database Schema)** | `mermaid` | Dev + BA | Nền phẳng pastel hiện đại, gọn gàng cho phân hệ 3–15 bảng |
| 8 | **Sequence Diagram** | `mermaid` | Dev + BA | Chuỗi gọi API thanh thoát, đánh số tự động `autonumber` |
| 9 | **Flowchart / Process** | `mermaid` | BA | Bẻ nhánh if/else tự do, đổi màu khối nhanh bằng CSS/style |
| 10 | **State Diagram** | `mermaid` | Dev | Trạng thái bo góc tròn, màu sắc hiện đại hơn nét vẽ thô |
| 11 | **Mind Map** | `mermaid` | BA + Dev | Phân rã tính năng nhanh, màu pastel chia nhánh trực quan |
| 12 | **Screen Flow (Interactive)** | `canvas` | Dev + BA | Định vị X,Y pixel-perfect, kéo thả trên `diagram_editor.py` |

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
Antigravity Office Studio features a full-fledged, high-aesthetic web application runnable locally or deployable to web servers:
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

### Included Customizations:
- `.agents/skills/doc-handler/SKILL.md`: Instructs AI agents on converting documents, fixing table splits, and formatting headings.
- `.agents/skills/technical-diagrams/SKILL.md`: Instructs AI agents on composing declarative JSON specs, calculating zero-overlap layouts, and routing lines.
- `.agents/rules/rule_decoupled_document_converter.md`: System invariants for document fidelity and diagram styling.
- `AGENTS.md`: Full AI agent operational guide.

### Sample AI Prompt:
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
tool/pdf_to_docx_converter/
├── .agents/                                     # Antigravity Agent Configuration
│   ├── rules/                                   # Universal Engineering Rules
│   └── skills/                                  # Agent Workflow Skills
├── specs/                                       # Declarative Diagram Specifications
│   ├── flow_*.json                              # Screen Flow JSON Specs
│   ├── master_erd_*.json                        # Master ERD Canvas & Connections Specs
│   └── sub_erd_*.json                           # Domain Sub-ERD Specs
├── tools/                                       # Generator & Document Patcher Scripts
│   ├── gen_canvas_erd.py                        # Master 4K ERD Generator
│   ├── gen_hub_spoke_flow.py                    # 25-Node Hub & Spoke Flow Generator
│   ├── patch_docx_*.py                          # Document Update & Diagram Injection Scripts
│   └── plantuml.jar                             # Local PlantUML Runtime
├── tests/                                       # Test Suites
│   ├── test_diagram_editor.py                   # Canvas Visual Editor Tests
│   ├── test_mermaid_suite.py                    # Mermaid Engine Tests
│   ├── test_plantuml_suite.py                   # PlantUML Engine & Dispatcher Tests
│   ├── test_edge_render.py                      # Headless Vector Rendering Tests
│   └── test_cardinality.py                      # Crow's Foot Geometry Tests
├── diagram_assets/                              # Rendered PNG & SVG Output Artifacts
├── ai_tools_cli.py                              # Unified CLI for Terminal & AI
├── spec_diagram_engine.py                       # Precision SVG + Chromium Engine
├── diagram_editor.py                            # Interactive Canvas Diagram Editor
├── mermaid_renderer.py                          # Mermaid CLI Wrapper & Themes
├── plantuml_renderer.py                         # PlantUML Engine & C4 Standard Library
├── smart_post_processor.py                      # OpenXML Word Table & TOC Repair
├── converter_engine.py                          # Multi-Format Pipeline Coordinator
├── markdown_converter.py                        # Decoupled MD + Style YAML Engine
├── docx_reader.py / docx_writer.py              # Low-Level OpenXML Word Handlers
├── xlsx_reader.py / xlsx_writer.py              # Excel Worksheet Utilities
├── gui.py / main.py                             # Desktop GUI Application
├── requirements.txt                             # Python Dependencies
├── pyproject.toml                               # Packaging Configuration
├── LICENSE                                      # MIT License
├── AGENTS.md                                    # Operational Guide for AI Agents
└── README.md                                    # Documentation
```

---

## 📄 License

This project is licensed under the [MIT License](LICENSE).
Feel free to use, modify, distribute, and integrate into your own workflows.
