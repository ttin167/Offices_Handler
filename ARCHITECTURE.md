# Architecture Guide — Antigravity Universal Office Studio

> **Phiên bản**: 4.6.0 (Enterprise Omni-Pipeline, Deep CDR, Cryptographic PAdES/PDF-A & AST Code-to-Diagram Architecture v4.6)  
> **Tiêu chuẩn**: Zero-Regression, 100% Backward Compatibility, Strict Layer Isolation, Clean-Room MIT Compliance  
> **Test Suite**: 323+ Automated Tests (100% Pass Rate) | **MCP Tools**: 28 Tools

---

## 1. Tổng Quan Kiến Trúc Phân Tầng (5-Layer Modular Architecture)

Hệ thống được tổ chức theo mô hình phân tầng chặt chẽ (Clean Architecture & Decoupled Modular Layers):

```
┌────────────────────────────────────────────────────────────────────────────────────────┐
│                              1. Presentation & Client Layer                            │
│   Web Studio (src/ - Vite ES6 + Glassmorphic) │ Desktop GUI & App Mode (ui/)           │
└───────────────────────────────────────────┬────────────────────────────────────────────┘
                                            │
                                            ▼
┌────────────────────────────────────────────────────────────────────────────────────────┐
│                              2. Interface & Protocol Layer                             │
│   Web REST API (server/app.py :8000)      │ Model Context Protocol (28 Tools)          │
│   SSE Task Manager (server/job_manager)   │ AI Unified CLI (cli/ai_tools + v46_cmds)   │
└───────────────────────────────────────────┬────────────────────────────────────────────┘
                                            │
                                            ▼
┌────────────────────────────────────────────────────────────────────────────────────────┐
│                            3. Core Document & Automation Engine                        │
│  core/docx/     │ core/pptx/     │ core/xlsx/     │ core/converter/ │ core/pipeline/   │
│  core/diagrams/ │ core/security/ │ core/yearbook/ │ core/env_doctor │                  │
└───────────────────────────────────────────┬────────────────────────────────────────────┘
                                            │
                                            ▼
┌────────────────────────────────────────────────────────────────────────────────────────┐
│                          4. Storage, Cache & Diagnostic Layer                          │
│   SQLite (history.db / db.py) │ In-Memory SSE Queue │ Unified QA Diagnostic (13 Gates) │
└───────────────────────────────────────────┬────────────────────────────────────────────┘
                                            │
                                            ▼
┌────────────────────────────────────────────────────────────────────────────────────────┐
│                      5. External Tooling & Headless Automation                         │
│   plantuml.jar (C4 stdlib) │ Headless Chromium/Edge │ LibreOffice Headless │ Tesseract │
└────────────────────────────────────────────────────────────────────────────────────────┘
```

---

## 2. Chi Tiết Trách Nhiệm Từng Tầng (Layer Responsibilities)

### 2.1. `core/` — Tầng Lõi Nghiệp Vụ Tài Liệu & Đồ Họa (Core Engines)
Tầng này chứa 100% logic thuật toán xử lý dữ liệu và tạo lập tài liệu văn phòng / đồ họa.
- **Quy tắc phụ thuộc**: Tuyệt đối không import từ `server/`, `ui/`, hoặc `cli/`. Chỉ nhận tham số thuần túy (Python dict, dataclass, file path) và trả về kết quả.
- **Các phân hệ cốt lõi**:
  - `core/docx/`:
    - `reader.py` / `writer.py`: Đọc và ghi tài liệu Word OpenXML, duy trì `<w:cantSplit/>`, `<w:tblHeader/>`, `<w:vAlign w:val="center"/>`.
    - `lists.py`: Động cơ đánh số tự động OpenXML (`numId` động, `startOverride=1`, bullet 3 cấp, checkbox `☐`/`☑`).
    - `math.py`: Kiến trúc toán học Dual-Mode (OpenXML Multi-Run & OMML `<m:oMath>`).
    - `diff_engine.py`: So sánh tài liệu pháp lý và AI Track Changes (`<w:ins>`, `<w:del>`, `<w:delText>`, `word/comments.xml`).
    - `mail_merge.py`: Studio trộn thư hàng loạt từ template Word + dataset Excel/JSON, hỗ trợ logic `{{IF}}` và xuất file `.zip`.
    - `nd30_preset.py`: Tự động áp dụng thể thức văn bản hành chính theo Nghị định 30/2020/NĐ-CP (lề in, quốc hiệu, tiêu ngữ, bảng 2 cột).
    - `vietnamese_sanitizer.py`: Chuẩn hóa typographic tiếng Việt (Unicode NFC, khoảng trắng dấu câu, En-dash `–`).
    - `charts.py`: Kết xuất vector SVG và PNG phân giải cao cho 5 loại biểu đồ, giới hạn lề in ấn $\le 14\text{cm}$.
  - `core/pptx/`:
    - `engine.py` / `writer.py`: Đọc, biên dịch slide Widescreen 16:9 (`PPTX_INV_01`), typography thuần đen `#000000` (`PPTX_INV_10`), run-level formatting (`PPTX_INV_09`), Adaptive Card Sizing (tự co giãn chiều cao và vị trí thẻ theo số dòng text thực tế).
    - `docx_to_pptx.py`: Bộ chuyển đổi thông minh từ Word sang slide thuyết trình 16:9; tích hợp Executive Deck Synthesizer cho tài liệu lớn $\ge 100$ trang (Target Deck Sizing 15-25 slides, Semantic Fluff Pruning, Quantitative KPI Prioritization, H3 minutiae folding, trần cứng $\le 25$ slides, Batch Progressive Previews 16 slides/lần).
    - `math.py`: Native DrawingML Math (`baseline="-25000"`, `baseline="30000"`, biến in nghiêng).
    - `charts.py`: Biểu đồ OpenXML Native (`<c:chart>`) nhúng gói bảng tính Excel ẩn cho phép "Edit Data" trực tiếp trên PowerPoint (`PPTX_INV_14`).
    - Business Layouts: Hỗ trợ 4 mẫu slide doanh nghiệp (KPI Summary Cards, Timeline Horizontal Roadmap, SWOT 2x2 Matrix, Process Chevron Steps).
  - `core/xlsx/`:
    - `reader.py` / `writer.py`: Đọc, mutate bảng tính Excel in-place, clone style prototype row (`ERR_XLSX_001`), dịch chuyển công thức động (`ERR_XLSX_002`), thiết lập trang in Print-Ready (`ERR_XLSX_008`), `safe_set_cell` tự động bảo vệ ô hợp nhất (`ERR_XLSX_004`), bảo toàn số 0 ở đầu (`ERR_XLSX_007`), kiểm soát 15 chữ số & parse ngày tiếng Việt / số âm kế toán (`ERR_XLSX_009`), khử ký tự rác XML & cắt tỉa 32k ký tự (`ERR_XLSX_010`), tiền tố `_xlfn.` cho hàm hiện đại (`ERR_XLSX_011`), hiển thị đường lưới `showGridLines`, `auto_filter.ref` và mở khóa chiều cao dòng dữ liệu (`ERR_XLSX_012`), bảo vệ chọn lọc công thức và cấu trúc (`ERR_XLSX_013`), định dạng có điều kiện tự động, ràng buộc nhập liệu Data Validation, biểu đồ sống In-Sheet Charts (`ERR_DIAG_003`), và bảng điều hướng Index Dashboard.
    - `streaming.py`: Động cơ xử lý dữ liệu lớn tiết kiệm bộ nhớ (`ERR_XLSX_014`) với `stream_read_xlsx` (`read_only=True`) và `stream_write_xlsx` (`write_only=True`), giữ mức tiêu thụ RAM luôn $\le 50\text{MB}$ khi xử lý tệp $> 10\text{MB}$ hoặc $> 30.000$ dòng trên Render 512MB.
    - `validator.py`: Kiểm định chất lượng 12 Cổng (Invariant E6).
    - `format_diff.py`: Bóc tách và so sánh định dạng giữa các sheet.
    - `formula_playbook.py`: Hệ thống Cẩm Nang 32+ Công Thức Excel Tinh Hoa & Live Demo (4 chiều: Tên hàm, Công thức mẫu, Dùng trong trường hợp nào, Tại sao lại chọn hàm đó của Excel). Tự động sinh sổ tính 2 sheets chuẩn Enterprise Navy: Sheet 1 tra cứu Playbook và Sheet 2 công thức động đang tính toán thực tế (`=SUMIFS`, `=COUNTIFS`, `=AVERAGEIFS`, `=XLOOKUP`, `=IFS`, `=TEXTJOIN`).
  - `core/converter/`:
    - `engine.py`: Dispatcher ma trận 5 chiều (PDF ↔ DOCX ↔ MD ↔ PPTX ↔ XLSX) và Decoupled Markdown parser.
    - `markdown.py`: Bộ phân tích và xuất bản Markdown tách rời kiểu dáng (Decoupled Markdown), tích hợp **Động cơ Mục Lục Tương Tác 3 Chiều** (Interactive TOC Engine: chuẩn hóa GFM slug, tự động đánh dấu `<w:bookmarkStart>` OpenXML $\le 40$ ký tự, liên kết hyperlink nội bộ `<w:hyperlink w:anchor="...">` phong cách Clean Publication Style, và tự động bảo toàn liên kết nội bộ `LINK_GOTO` trên PDF).
    - `table_extractor.py`: Semantic Table Extractor với Matrix Post-Processor; ưu tiên chiến lược `lines_strict` vector borders, thuật toán cắt bỏ cột ma `_prune_ghost_columns`, ghép nối bảng đa trang liên tục `_stitch_multipage_tables`, tái tạo ô gộp tiêu đề 2 tầng `_detect_table_header_structure`, cố định dòng tiêu đề `freeze_panes = 'A3'/'A2'`, lọc bảng hành chính NĐ30, xử lý an toàn ô gộp `<w:gridSpan>` và `<w:vMerge>` (`ERR_XLSX_004`), giữ nguyên số 0 ở đầu cho mã định danh `@` (`ERR_XLSX_007`), thiết lập in ấn A4 ngang (`ERR_XLSX_008`).
    - `ocr_pipeline.py`: Pipeline bóc tách PDF/ảnh scan tiếng Việt (`vie+eng`), bộ lọc mộc đỏ HSV (`ERR_CONV_005`), nắn thẳng ảnh scan tự động (`deskew`).
    - `pdf_renderer.py`: Bộ chuyển đổi PDF phổ quát (`UniversalPdfRenderer`) với chuỗi ưu tiên: Win32 COM ➔ LibreOffice Headless ➔ PyMuPDF.
    - `pdf_toolkit.py`: Phân hệ PDF Pro Studio & iLovePDF Killer Features (True Redaction `ERR_PDF_001`, AES-256 Encryption, Dấu giáp lai đa trang overlap 8% `ERR_PDF_002`, Chuẩn hóa A4 vector `ERR_PDF_003`, Bóc bảng sang Excel `ERR_XLSX_007`, Tổ chức trang trực quan kèm chèn trang trắng A4, Chuyển đổi 2 chiều Ảnh ↔ PDF, Ký tên điện tử & đóng dấu bù góc xoay `page.derotation_matrix`, Làm phẳng 2 chế độ Vector Bake & Strict Raster kèm dọn sạch metadata ẩn `doc.scrub`).
  - `core/yearbook/`:
    - `layout_engine.py`: Động cơ dàn trang thông minh phân tích EXIF (`ERR_EXIF_001`), phân loại tỷ lệ ảnh (Landscape, Portrait, Square), tự động dàn trang 9 mẫu layout chuẩn tạp chí (`cover_hero`, `editorial_quote`, `symmetric_duo`, `magazine_grid_4`, `polaroid_scrapbook`, `mosaic_gallery_6`, `classroom_grid_8`, `megagrid_12`, `panoramic_spread`), thuật toán gom cụm mật độ cao cho 50–100+ ảnh, và sinh album mẫu Zero-Config Mock Album.
    - `pdf_press_exporter.py`: Động cơ xuất xưởng in ấn công nghiệp 300+ DPI với lề xén 3mm Bleed, 4 cặp vạch cắt Crop Marks, bù trừ rãnh gáy sách Gutter, tự động nhúng font hệ thống Unicode tiếng Việt (`ERR_PDF_FONT_001`), cắt ảnh object-fit cover với top-bias 35% giữ nguyên gương mặt và mũ cử nhân (`ERR_IMG_CROP_001`), sinh bản PDF mở phẳng (Spread View) và thumbnail PNG.
    - `handoff_packager.py`: Đóng gói hồ sơ bàn giao thợ thiết kế 5-trong-1 (.zip) với cơ chế stream đĩa trực tiếp (`ERR_RAM_001`), sinh bản Word DOCX Brief tuân thủ nghiêm ngặt các bất biến OpenXML bảng biểu (`<w:cantSplit/>`, `<w:tblHeader/>`, `<w:vAlign w:val="center"/>`, The Last Paragraph Rule).
  - `core/security/`:
    - `document_auditor.py`: Kiểm định an ninh tài liệu tĩnh dựa trên bộ kiểm tra phòng thủ Cloudflare Security Audit Harness, phân định 3 trạng thái phản biện đối nghịch (`confirmed`, `needs_validation`, `rejected`) cho 5 bất biến: `ERR_SEC_001` (Macro streams & PDF `/Launch`), `ERR_SEC_002` (Formula / DDE Injection), `ERR_SEC_003` (OLE MZ headers & PE binaries), `ERR_SEC_004` (PII author metadata & `veryHidden` sheets), `ERR_SEC_005` (Zip Slip `../` traversal & XXE DTD entities).
    - `sanitizer.py`: Tẩy rửa và vô hiệu hóa mối đe dọa không phá hủy (Non-destructive Defuser & Deep CDR `ERR_SEC_006` + Smart Vietnamese PII Redactor `ERR_SEC_007`): Thêm tiền tố nháy đơn `'` cho công thức nguy hiểm, bóc gỡ `vbaProject.bin` & OLE PE (`MZ`), tự động chuyển đổi `macroEnabled.main+xml` sang chuẩn OpenXML, dọn sạch `.rels`, và che mờ thông minh CCCD/SĐT/STK/MST/Email không phạm vào số liệu tài chính BCTC.
  - `core/pipeline/`:
    - `engine.py`: Động cơ đường ống công việc khai báo đa bước (`WorkflowPipeline`, `run_pipeline`, `run_bulk_directory_pipeline`) tuân thủ `ERR_PIPE_001` (giới hạn `MAX_JOB_WORKERS=2` bảo vệ container 512MB RAM và dọn sạch 100% tệp trung gian `step_*_temp.*` trong khối `try...finally`).
  - `core/env_doctor/`:
    - `run_doctor.py`: Tự động quét kiểm tra 12 thành phần môi trường (Python, LibreOffice, Tesseract, Chromium, PlantUML, font chữ) và xuất báo cáo JSON / Terminal kèm mã lệnh khắc phục 1-click.
  - `core/diagrams/`:
    - `bridge.py`: Điều phối sơ đồ thống nhất.
    - `mxgraph.py`: Bộ sinh sơ đồ cấu trúc Draw.io XML native (`MX_INV_01`..`20`) và tự động nhúng vào `.docx` / `.pptx`.
    - `code_to_diagram.py`: Bộ quét mã nguồn tĩnh Python AST & SQL/Prisma DDL sinh sơ đồ Class/ERD tự động kèm bộ xuất Vector SVG thuần Python (`export_mxgraph_to_svg`).
    - `editor.py`: Trình biên tập Canvas tương tác 2 chiều.
    - `legacy/`: Các renderer kế thừa (`canvas`, `plantuml`, `mermaid`).

### 2.2. `server/` — Tầng Giao Thức & Cung Cấp Dịch Vụ (Web API, MCP, SSE)
- `app.py`: REST API Server nhẹ nhàng dựa trên `http.server`, phục vụ Web Studio cổng `8000`. Cung cấp endpoints cho conversion, history, mail merge, ocr, diff, math, static assets, trọn bộ PDF Pro Endpoints, Atomic Office Endpoints, Formula Playbook Endpoints, và trọn bộ Digital Yearbook Endpoints (`/api/yearbook/mock-album`, `/api/yearbook/presets`, `/api/yearbook/auto-layout`, `/api/yearbook/render-preview`, `/api/yearbook/export`).
- `mcp_server.py`: Model Context Protocol Server (STDIO JSON-RPC 2.0) tương thích với Claude Desktop, Cursor, và Antigravity IDE. Đăng ký **23 tools** chuẩn:
  - Macro Tools (11): `env_doctor`, `doc_convert`, `doc_inspect`, `doc_insert_diagram`, `docx_patch`, `docx_apply_preset`, `docx_sanitize_vietnamese`, `pptx_synthesize`, `diagram_render`, `doc_mail_merge`, `doc_redline_diff`, `doc_ocr_scan`, `pdf_quick_tools`.
  - Atomic Office Tools (8): `xlsx_set_cell`, `xlsx_set_range`, `xlsx_insert_rows`, `xlsx_find_anchor`, `docx_inspect_placeholders`, `docx_replace_text`, `docx_insert_paragraph`, `docx_insert_table`.
  - Formula Playbook Tool (1): `xlsx_formula_helper`.
  - Yearbook Studio Tools (2): `yearbook_auto_layout`, `yearbook_export`.
  - Document Security Tool (1): `doc_security_audit`.
- `job_manager.py`: Server-Sent Events (SSE) background task manager và job queue, cho phép theo dõi tiến trình OCR, mail merge, render và xuất bản kỷ yếu theo thời gian thực.
- `db.py`: Zero-config SQLite database (`history.db`) lưu trữ lịch sử chuyển đổi và tải tệp.

### 2.3. `cli/` — Tầng Giao Diện Dòng Lệnh (CLI Layer)
`cli/ai_tools.py` và facade tại root `ai_tools_cli.py` cung cấp 33 subcommands hoàn chỉnh cho AI Agents và terminal users:
- `doctor`: Kiểm tra sức khỏe môi trường.
- `convert`: Chuyển đổi phổ quát 5 định dạng (PDF, DOCX, MD, PPTX, XLSX).
- `md-toc`: Tạo, xem trước và tự động đồng bộ hóa mục lục tương tác trực tiếp trong tệp Markdown.
- `docx-read`, `docx-write`, `docx-inject`, `docx-patch`, `docx-diff`: Quản lý tài liệu Word.
- `docx-tokens`, `docx-replace`: Atomic Word CLI tools (quét biến & thay thế text bảo tồn OpenXML run).
- `xlsx-set`, `xlsx-insert-rows`: Atomic Excel CLI tools (sửa ô hợp nhất & mở rộng dòng kèm dịch công thức).
- `xlsx-formulas`: Cẩm nang tra cứu và xuất Excel Formula Playbook 2 sheets.
- `yearbook-layout`, `yearbook-build`: Tự động dàn trang photobook và đóng gói ZIP bàn giao thợ thiết kế / nhà in.
- `doc-security`: Kiểm tra và tẩy rửa an ninh tài liệu, formula injection, OLE PE binary, và PII (Cloudflare Harness).
- `docx-mail-merge`: Trộn thư hàng loạt Word + Excel/JSON.
- `docx-apply-preset`, `docx-sanitize-vietnamese`: Chuẩn hóa thể thức hành chính & chính tả.
- `pptx-build`, `pptx-inspect`, `pptx-preview`, `pptx-insert-diagram`, `docx-to-pptx`, `pdf-to-pptx`: Quản lý thuyết trình PowerPoint.
- `pdf-tools`: Bộ 16 công cụ PDF Pro Studio (`merge`, `split`, `rotate`, `clean`, `watermark`, `compress`, `redact`, `protect`, `normalize-a4`, `giap-lai`, `extract-tables`, `images-to-pdf`, `export-images-zip`, `sign`, `flatten`, `reorganize`).
- `ocr-scan`: Bóc tách tài liệu scan và loại bỏ dấu mộc đỏ.
- `diagram-render`, `plantuml-render`, `mermaid-render`, `spec-render`, `diagram-editor`: Sơ đồ kỹ thuật.
- `qa-audit`: Master QA diagnostic and repair pipeline 13 Cổng (`tools/unified_qa_diagnostic.py`).

### 2.4. `ui/` & `src/` — Tầng Giao Diện Người Dùng (Desktop GUI & Web Studio)
- `ui/`: Desktop GUI Tkinter độc lập (`main.py`) và Native WebView App Mode (`ui/webview_launcher.py`) hỗ trợ chạy offline không cần trình duyệt.
- `src/`: Web Studio trên nền Vite ES6 modules và Vanilla CSS Glassmorphism:
  - `src/main.js` (~295 dòng): Orchestrator nhẹ quản lý theme, toast notifications, health polling, tab routing và khởi tạo các studio modules.
  - `src/modules/`: Tách rời hoàn toàn theo từng phân hệ chuyên biệt: `presentation_studio.js`, `spreadsheet_studio.js`, `document_hub.js`, `diagram_studio.js`, `math_studio.js`, `formula_playbook.js`, `history_drawer.js`, `constants.js`, và `yearbook_studio.js` (optional local-only module qua `import.meta.glob`).


### 2.5. `tools/` — Tầng Tiện Ích QA & Giao Nộp (Diagnostics & Deliverables)
- `tools/unified_qa_diagnostic.py`: Động cơ chẩn đoán toàn diện 13 Cổng cho DOCX, XLSX, PPTX, Draw.io, và ảnh scan với khả năng tự sửa lỗi (Safe Non-destructive Repair).
- `tools/format_diff_excel.py`: Tiện ích so khớp định dạng 12 cổng bảng tính Excel.
- `tools/bump_version.py` & `tools/sync_core_docs.py`: Tự động đồng bộ 10 điểm neo phiên bản (`ERR_VER_001`) và kiểm định đối xứng ngữ nghĩa bộ tài liệu lõi (`ERR_DOC_SYNC_001`).
- `tools/gen_*.py`: Các bộ sinh sơ đồ kỹ thuật chuẩn (`gen_canvas_erd.py`, `gen_hub_spoke_flow.py`, `gen_master_erd.py`, `gen_usecase_diagram.py`, `gen_illustration_assets.py`).
- `tools/deliverables/`: Lưu trữ riêng các scripts xuất báo cáo/đồ án cá nhân (được loại trừ trong `.gitignore` để giữ kho mã nguồn chính luôn gọn sạch).

---

## 3. Tiêu Chuẩn Clean-Room MIT & Tương Thích Linux/Docker (Zero-COM)

Dự án được thiết kế theo chuẩn Clean-Room độc lập 100% với hệ điều hành:
1. **Chuỗi Fallback Zero-COM**:
   - Trên Windows: Ưu tiên Win32 COM nếu có MS Office và không bị khóa tiến trình.
   - Trên Linux / macOS / Docker: Tự động chuyển tiếp sang LibreOffice Headless (`soffice --headless --convert-to pdf`) hoặc thư viện Python thuần (`PyMuPDF`, `openpyxl`, `python-docx`, `python-pptx`).
2. **Docker Container Sẵn Sàng**:
   - `Dockerfile` & `docker-compose.yml` trên nền Debian Bookworm đã cài sẵn đầy đủ font chữ Unicode (Liberation, DejaVu), OpenJDK JRE, LibreOffice, Tesseract OCR (`vie+eng`), và Chromium.

---

## 4. Ma Trận Bất Biến & Cổng Kiểm Định Chất Lượng (Invariant Quality Gates)

### 4.1. Word Document Invariants (ERR_DOCX)
- `ERR_DOCX_001` (**The Last Paragraph Rule**): Mọi cell bảng (`<w:tc>`) bắt buộc kết thúc bằng tối thiểu một thẻ `<w:p>`.
- `ERR_DOCX_002` (**Run Text Overwrite**): Cập nhật văn bản qua `run.text`, không gán `cell.text = "..."`.
- `ERR_DOCX_003` (**Multi-Page Table Flags**): Bảng $\ge 10$ hàng bắt buộc có `<w:cantSplit/>` và `<w:tblHeader/>`.
- `ERR_DOCX_005` (**Printable Margin Overflow**): Khóa tỷ lệ ảnh; chiều rộng $\le$ chiều rộng vùng in ($\le 14\text{cm}$ trên A4).
- `ERR_DOCX_008` (**Valid Track Changes XML**): Thẻ `<w:ins>` / `<w:del>` phải có `w:id` dương duy nhất, `w:author`, `w:date` ISO-8601, và dùng `<w:delText>` cho văn bản xóa.

### 4.2. Excel Spreadsheet Invariants (ERR_XLSX)
- `ERR_XLSX_001` (**Prototype Row Style Cloning**): Clone 100% style từ hàng mẫu template sang hàng mới.
- `ERR_XLSX_002` (**Dynamic Formula Shifting**): Dịch chuyển công thức động qua Regex khi chèn thêm hàng.
- `ERR_XLSX_004` (**Safe Merged-Cell Handling**): Chỉ gán giá trị vào Top-Left; đồng bộ border toàn dải ô gộp.
- `ERR_XLSX_005` (**Freeze Panes & Dynamic Height**): Cố định tiêu đề bảng và tính chiều cao hàng theo số dòng text wrap.
- `ERR_XLSX_006` (**DrawingML Preservation**): Giữ nguyên shape, biểu đồ và logo công ty từ template gốc.
- `ERR_XLSX_007` (**Identifier Number Format**): Cột mã định danh (ID, Code, MST) bắt buộc gán `number_format = '@'`.
- `ERR_XLSX_008` (**Print-Ready Setup**): Khóa trang in A4, `fitToPage=True`, lề chuẩn, căn giữa trang ngang.

### 4.3. PowerPoint Presentation Invariants (PPTX_INV)
- `PPTX_INV_01` (**Widescreen 16:9 Master**): Slide xuất tỷ lệ chuẩn 16:9 (13.333 x 7.5 in / 1280 x 720 pt).
- `PPTX_INV_02` (**Aspect-Ratio Lock**): Hình ảnh và sơ đồ khóa tỷ lệ khung hình, không bị méo tỷ lệ.
- `PPTX_INV_04` (**Zero Text Overflow & Word Wrap**): Bắt buộc `word_wrap = True` trên mọi text frame có nội dung.
- `PPTX_INV_06` (**Clean Native Tables**): Dùng `<p:graphicFrame>` bảng chuẩn, header fill rõ ràng.
- `PPTX_INV_07` (**Windows File Lock Guard**): Xử lý an toàn khi tệp PPTX đang mở trong PowerPoint.
- `PPTX_INV_09` (**Run-Level Formatting Supremacy**): CẤM đặt bold/fill trên `<a:defRPr>`. Mọi format phải đặt ở `<a:rPr>`.
- `PPTX_INV_10` (**Default Pure Black Typography**): Mọi caption, subtitle, card header mặc định dùng đen thuần `#000000`.
- `PPTX_INV_14` (**Native OpenXML Charts & Embedded Excel**): Biểu đồ dùng thẻ `<c:chart>` nhúng workbook Excel ẩn.

### 4.4. Conversion & Scan Invariants (ERR_CONV)
- `ERR_CONV_004` (**Zombie Lock Files Cleanup**): Quét và dọn sạch các tệp lock rác `.~lock.*` và `~$*`.
- `ERR_CONV_005` (**Scan & Red-Stamp Recovery**): Lọc dấu mộc đỏ qua không gian màu HSV và nắn góc nghiêng ảnh scan.

### 4.5. Draw.io & Diagram Invariants (MX_INV)
- `MX_INV_01` (**Pure Native Hierarchy**): Sơ đồ Draw.io là các thẻ `<mxCell>` native, không bọc `<UserObject>` chứa raw code.
- `MX_INV_02` (**Minimalist & Well-Formed XML**): XML sạch, không chứa metadata rác, đóng thẻ 100%.
- `MX_INV_08` (**Collision-Free Labels & Masks**): Mọi nhãn đường nối line phải có `labelBackgroundColor=#ffffff` chống đè chữ.

---

## 5. Bất Biến Điểm Nhập & Tương Thích Ngược (Entrypoint & Compatibility Guarantees)

1. **6 Entrypoints Chính Thức Tại Thư Mục Gốc**:
   Thư mục gốc chỉ giữ lại đúng 6 file entrypoint chính thức (`main.py`, `ai_tools_cli.py`, `web_api_server.py`, `mcp_server.py`, `spec_diagram_engine.py`, `diagram_editor.py`) cùng `__init__.py`. Toàn bộ các mô-đun nội bộ và bài kiểm thử trong `tests/` import trực tiếp từ cấu trúc phân tầng `core.*`, `server.*`, `cli.*`, `ui.*`.
2. **100% Pass Rate Toàn Bộ Test Suite**:
   Tất cả 323+ bài kiểm thử tự động trong `tests/` phải luôn luôn PASS 100% trước khi phát hành.
3. **Giữ Vững CLI Contract**:
   Tất cả cú pháp lệnh trong `AGENTS.md` và `pyproject.toml` tiếp tục hoạt động chính xác không thay đổi.
