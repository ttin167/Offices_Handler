# Architecture Guide — Antigravity Universal Office Studio

> **Phiên bản**: 5.0.0 (Pure Office IT Studio — Excel Power Suite, Word NĐ30 & `vn_words`, PDF Pro 17 Tools, AST Offline Formula Engine, Vector Print & Universal FTS5 RAG Copilot)  
> **Tiêu chuẩn**: Zero-Regression, 100% Backward Compatibility, Strict Layer Isolation, Clean-Room MIT Compliance  
> **Test Suite**: Automated Tests (100% Pass Rate) | **MCP Tools**: 43 Tools

---

## 1. Tổng Quan Kiến Trúc Phân Tầng (5-Layer Modular Architecture)

Hệ thống được tổ chức theo mô hình phân tầng chặt chẽ (Clean Architecture & Decoupled Modular Layers):

```
┌────────────────────────────────────────────────────────────────────────────────────────┐
│                              1. Presentation & Client Layer                            │
│   Web Studio (src/ - React 18 + TS + Tailwind v3) │ Desktop GUI & App Mode (ui/)       │
└───────────────────────────────────────────┬────────────────────────────────────────────┘
                                            │
                                            ▼
┌────────────────────────────────────────────────────────────────────────────────────────┐
│                              2. Interface & Protocol Layer                             │
│   Web REST API (server/app.py :8000)      │ Model Context Protocol (43 Tools)          │
│   SSE Task Manager (server/job_manager)   │ AI Unified CLI (cli/ai_tools + office+v46) │
└───────────────────────────────────────────┬────────────────────────────────────────────┘
                                            │
                                            ▼
┌────────────────────────────────────────────────────────────────────────────────────────┐
│                            3. Core Document & Automation Engine                        │
│  core/docx/     │ core/pptx/     │ core/xlsx/     │ core/converter/ │ core/pipeline/   │
│  core/diagrams/ │ core/security/ │ core/copilot/  │ core/env_doctor │ core/maintenance │
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
    - `design_system.py`: Centralized Enterprise Document Design System & Premium Styler:
      - Đa Theme Palette (`modern_tech` [Slate/Indigo], `executive_classic` [Deep Navy/Burgundy], `nghidinh30` [Decree 30 Standard]).
      - Executive Cover Page Builder với thanh dải Bauhaus accent và bảng thông số metadata card.
      - Dynamic Table of Contents tự động (`<w:fldSimple w:instr="TOC ..."/>`, cập nhật qua F9).
      - Header & Footer đồng bộ 2 bên (tiêu đề tài liệu bên phải, phân loại bảo mật bên trái, `Trang {PAGE} / {NUMPAGES}` bên phải).
      - Callout Boxes viền màu pastel (`NOTE`, `TIP`, `WARNING`, `IMPORTANT`, `CAUTION`) bảo toàn invariants OpenXML.
      - Fenced Code Block Cards bo góc có nhãn ngôn ngữ và font Consolas.
      - Bảng biểu Modern Clean Studio (padding 6pt/8pt, viền mảnh `#E2E8F0`, tự động căn lề theo kiểu dữ liệu cột).
      - Token Formatters động (`{{KEY:VND}}`, `{{KEY:USD}}`, `{{KEY:VN_WORDS}}`, `{{KEY:WORDS_PLAIN}}`, `{{KEY:DATE}}`, `{{KEY:DATE_VN}}`, `{{KEY:UPPER}}`, `{{KEY:LOWER}}`, `{{KEY:TITLE}}`, `{{KEY:INT}}`, `{{KEY:NUM}}`) và bộ chuyển đổi đọc số tiền thành chữ tiếng Việt `number_to_vietnamese_words`.
    - `reader.py` / `writer.py`: Đọc và ghi tài liệu Word OpenXML, duy trì `<w:cantSplit/>`, `<w:tblHeader/>`, `<w:vAlign w:val="center"/>`; hỗ trợ vá hình ảnh động `{{IMAGE:key}}` và ngắt dòng OpenXML `<w:br/>` an toàn trong `patch_docx_template`. Bảo vệ xung đột tệp Windows bằng `safe_save_docx`.
    - `toolkit.py`: Bộ tiện ích Word Mastery chuẩn hóa văn bản theo Nghị định 30/2020/NĐ-CP & Executive (`polish_docx_document` kèm đánh số trang `Page X/Y` & giới hạn ảnh lề in `ERR_DOCX_005`), ghép nối nhiều tệp Word bảo toàn ảnh/bảng (`merge_docx_files`), tách chương theo Heading ra gói ZIP (`split_docx_by_heading`), và đóng dấu chéo VML Watermark + Header/Footer (`apply_docx_watermark_and_hf`).
    - `lists.py`: Động cơ đánh số tự động OpenXML (`numId` động, `startOverride=1`, bullet 3 cấp, checkbox `☐`/`☑`).
    - `math.py`: Kiến trúc toán học Dual-Mode (OpenXML Multi-Run & OMML `<m:oMath>`).
    - `diff_engine.py`: So sánh tài liệu pháp lý và AI Track Changes (`<w:ins>`, `<w:del>`, `<w:delText>`, `word/comments.xml`).
    - `mail_merge.py`: Studio trộn thư hàng loạt từ template Word + dataset Excel/JSON, hỗ trợ logic `{{IF}}`, lặp dòng bảng danh sách con lồng nhau `{{#TABLE:Items}}...{{/TABLE}}` (`_process_nested_table_blocks`) và xuất file `.zip`.
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
    - `streaming.py`: Động cơ xử lý dữ liệu lớn tiết kiệm bộ nhớ (`ERR_XLSX_014`) với `stream_read_xlsx` (`read_only=True`) và `stream_write_xlsx` (`write_only=True`), giữ mức tiêu thụ RAM luôn $\le 50\text{MB}$ khi xử lý tệp $> 10\text{MB}$ hoặc $> 30.000$ dòng.
    - `validator.py`: Kiểm định chất lượng 12 Cổng (Invariant E6).
    - `format_diff.py`: Bóc tách và so sánh định dạng giữa các sheet.
    - `formula_playbook.py`: Hệ thống Cẩm Nang 57+ Công Thức Excel Tinh Hoa & Live Demo (4 chiều: Tên hàm, Công thức mẫu, Dùng trong trường hợp nào, Tại sao lại chọn hàm đó của Excel). Tự động sinh sổ tính 2 sheets chuẩn Enterprise Navy: Sheet 1 tra cứu Playbook và Sheet 2 công thức động đang tính toán thực tế (`=SUMIFS`, `=COUNTIFS`, `=AVERAGEIFS`, `=XLOOKUP`, `=IFS`, `=TEXTJOIN`).
    - `power_tools.py`: Bộ 6 công cụ Excel Power Suite thuần OpenPyXL: Làm sạch & chuẩn hóa dữ liệu VN (trim, chuẩn ngày dd/mm/yyyy, giữ số 0 đầu `@` SĐT/CCCD, khử trùng lặp `clean_and_standardize_workbook`), Tách/Gộp cột & tách họ tên tiếng Việt (`transform_workbook_columns`), Tổng hợp Pivot & Biểu đồ thực live `SUMIFS`/`COUNTIFS` (`build_pivot_summary_and_chart`), Nối bảng XLOOKUP không cần code (`xlookup_join_workbooks`), Tách bảng theo giá trị cột ra ZIP (`split_workbook_by_column`), và So sánh Excel Diff tô màu Amber/Emerald/Rose + `_Diff_Summary` (`diff_workbooks_visual`).
    - `tracker.py`: Động cơ Sổ Theo Dõi Công Việc & Báo Cáo Ngày (`OfficeWorkTracker`), tương thích hoàn toàn cấu trúc sheet/cột của `project_tracking.xlsx` và `task_tracking.xlsx` (`FSD_Mockup_Log` & Daily sheets).
    - `formula_evaluator.py`: Động cơ tính toán công thức ngoại tuyến dựa trên cây cú pháp trừu tượng (`ASTFormulaEvaluator`, `SafeFormulaASTTransformer`), hỗ trợ 30+ hàm cốt lõi (`SUM`, `AVERAGE`, `COUNT`, `COUNTA`, `COUNTIF`, `COUNTIFS`, `SUMIF`, `SUMIFS`, `AVERAGEIF`, `AVERAGEIFS`, `IF`, `IFS`, `IFERROR`, `XLOOKUP`, `VLOOKUP`, `INDEX`, `MATCH`, `TEXTJOIN`, ...) và nạp trực tiếp giá trị tính toán vào thẻ `<v>...</v>` OpenXML (`populate_cached_values`).
  - `core/converter/`:
    - `com_watchdog.py`: Bộ giám sát và cô lập tiến trình Win32 COM (`ComProcessWatchdog`) với cơ chế timeout 20s, tự động giải phóng `CoUninitialize` và thu hồi các tiến trình `WINWORD.EXE` / `POWERPNT.EXE` bị treo.
    - `engine.py`: Dispatcher ma trận 5 chiều (PDF ↔ DOCX ↔ MD ↔ PPTX ↔ XLSX) và Decoupled Markdown parser.
    - `markdown.py`: Bộ phân tích và xuất bản Markdown tách rời kiểu dáng (Decoupled Markdown), tích hợp **Động cơ Mục Lục Tương Tác 3 Chiều** (Interactive TOC Engine: chuẩn hóa GFM slug, tự động đánh dấu `<w:bookmarkStart>` OpenXML $\le 40$ ký tự, liên kết hyperlink nội bộ `<w:hyperlink w:anchor="...">` phong cách Clean Publication Style, và tự động bảo toàn liên kết nội bộ `LINK_GOTO` trên PDF).
    - `table_extractor.py`: Semantic Table Extractor với Matrix Post-Processor; ưu tiên chiến lược `lines_strict` vector borders, thuật toán cắt bỏ cột ma `_prune_ghost_columns`, ghép nối bảng đa trang liên tục `_stitch_multipage_tables`, tái tạo ô gộp tiêu đề 2 tầng `_detect_table_header_structure`, cố định dòng tiêu đề `freeze_panes = 'A3'/'A2'`, lọc bảng hành chính NĐ30, xử lý an toàn ô gộp `<w:gridSpan>` và `<w:vMerge>` (`ERR_XLSX_004`), giữ nguyên số 0 ở đầu cho mã định danh `@` (`ERR_XLSX_007`), thiết lập in ấn A4 ngang (`ERR_XLSX_008`).
    - `ocr_pipeline.py`: Pipeline bóc tách PDF/ảnh scan tiếng Việt (`vie+eng`), bộ lọc mộc đỏ HSV (`ERR_CONV_005`), nắn thẳng ảnh scan tự động (`deskew`).
    - `pdf_renderer.py`: Bộ chuyển đổi PDF phổ quát (`UniversalPdfRenderer`) với chuỗi ưu tiên: Win32 COM ➔ LibreOffice Headless ➔ PyMuPDF.
    - `pdf_toolkit.py`: Phân hệ PDF Pro Studio & iLovePDF Killer Features (True Redaction `ERR_PDF_001`, AES-256 Encryption, Dấu giáp lai đa trang overlap 8% `ERR_PDF_002`, Chuẩn hóa A4 vector `ERR_PDF_003`, Bóc bảng sang Excel `ERR_XLSX_007`, Tổ chức trang trực quan kèm chèn trang trắng A4, Chuyển đổi 2 chiều Ảnh ↔ PDF, Ký tên điện tử & đóng dấu bù góc xoay `page.derotation_matrix`, Làm phẳng 2 chế độ Vector Bake & Strict Raster kèm dọn sạch metadata ẩn `doc.scrub`, Quét và điền biểu mẫu tương tác AcroForms `inspect_pdf_form_fields` & `fill_pdf_form_fields`).
  - `core/copilot/` (Universal FTS5 RAG & Hybrid AI Copilot Engine):
    - `copilot_provider.py`: Kết nối đa hình Hybrid LLM (`BaseCopilotProvider`, `LocalOllamaProvider` tại `localhost:11434`, `GoogleGeminiProvider` ngữ cảnh 1M+ tokens, và `CopilotOrchestrator`) kèm tổng hợp RAG ngoại tuyến và sinh Thẻ Hành Động Tin học Văn phòng (`CopilotActionCard`).
    - `universal_rag.py`: Động cơ lập chỉ mục và truy xuất ngữ nghĩa đa định dạng (`UniversalRAGEngine`) dùng SQLite FTS5 (`unicode61 remove_diacritics 2`, chế độ WAL) cho các tệp PDF, DOCX, XLSX, CSV, JSON, và Markdown.
  - `core/printing/` (Enterprise Printing Engine):
    - `printer_manager.py`: Bộ điều phối máy in doanh nghiệp, giao tiếp trực tiếp Windows Spooler API qua `win32print`, hỗ trợ truy vấn trạng thái phần cứng và hàng đợi (`get_printer_status`) phát hiện sớm hết giấy (`JOB_STATUS_PAPEROUT`), kẹt giấy (`PRINTER_STATUS_PAPER_JAM`), ngoại tuyến (`OFFLINE`), cấu hình phần cứng `DEVMODE` (chọn khay giấy `tray_id`, in 2 mặt tiêu chuẩn `duplex`: `vertical` lật cạnh dài, `horizontal` lật cạnh ngắn, `simplex` 1 mặt) kèm chốt chặn **G2 (`DeviceCapabilities DC_BINS` / `DC_DUPLEX` pre-check)** chống lỗi trên máy in ảo, tự động phát hiện danh sách máy in mạng/nội bộ, hỗ trợ in ngầm (Silent Print) không cần hộp thoại UI, cơ chế fallback Edge Headless và Virtual Spool đệm PDF trên môi trường Linux/nội bộ cục bộ.
    - `print_optimizer.py`: Bộ tiền xử lý tài liệu in ấn bằng PyMuPDF (`fitz`): Phân tích cú pháp dải trang thông minh (`parse_page_ranges`) hỗ trợ biểu thức `1-5, 8, 11-20` và trích xuất trang (`doc.select`), chuẩn hóa khổ giấy đích (`A4`, `Letter`, `A3`) và căn chỉnh lề an toàn (`fit` thu vừa vùng in máy in chống mất mép viền bảng biểu `ERR_PRINT_001` vs `actual` giữ 100% kích thước gốc), chuyển đổi thang xám Vector nguyên bản (`_apply_vector_grayscale` chuyển toán tử màu sang `/DeviceGray` bảo toàn 100% font chữ vector), bố cục ghép trang đa trang N-Up (2-Up / 4-Up), sắp xếp thứ tự trang in đóng sách kèm bù trượt lề gáy sách (`_apply_booklet_imposition` Shingling Creep Compensation $\Delta x = k \times \text{thickness\_pt}$), dấu thủy ấn bảo mật (Confidential Watermark), và chốt chặn an toàn lề in chống mất chữ `ERR_PRINT_001` (Margin Safety Guard $\ge 5\text{mm}$).
  - `core/security/`:
    - `document_auditor.py`: Kiểm định an ninh tài liệu tĩnh dựa trên bộ kiểm tra phòng thủ Cloudflare Security Audit Harness, phân định 3 trạng thái phản biện đối nghịch (`confirmed`, `needs_validation`, `rejected`) cho 5 bất biến: `ERR_SEC_001` (Macro streams & PDF `/Launch`), `ERR_SEC_002` (Formula / DDE Injection), `ERR_SEC_003` (OLE MZ headers & PE binaries), `ERR_SEC_004` (PII author metadata & `veryHidden` sheets), `ERR_SEC_005` (Zip Slip `../` traversal & XXE DTD entities).
    - `sanitizer.py`: Tẩy rửa và vô hiệu hóa mối đe dọa không phá hủy (Non-destructive Defuser & Deep CDR `ERR_SEC_006` + Smart Vietnamese PII Redactor `ERR_SEC_007`): Thêm tiền tố nháy đơn `'` cho công thức nguy hiểm, bóc gỡ `vbaProject.bin` & OLE PE (`MZ`), tự động chuyển đổi `macroEnabled.main+xml` sang chuẩn OpenXML, dọn sạch `.rels`, và che mờ thông minh CCCD/SĐT/STK/MST/Email không phạm vào số liệu tài chính BCTC.
  - `core/pipeline/`:
    - `engine.py`: Động cơ đường ống công việc khai báo đa bước (`WorkflowPipeline`, `run_pipeline`, `run_bulk_directory_pipeline`) tuân thủ `ERR_PIPE_001` (giới hạn `MAX_JOB_WORKERS=2` kiểm soát RAM máy trạm và dọn sạch 100% tệp trung gian `step_*_temp.*` trong khối `try...finally`).
  - `core/env_doctor/`:
    - `run_doctor.py`: Tự động quét kiểm tra 12 thành phần môi trường (Python, LibreOffice, Tesseract, Chromium, PlantUML, font chữ) và xuất báo cáo JSON / Terminal kèm mã lệnh khắc phục 1-click.
  - `core/diagrams/`:
    - `bridge.py`: Điều phối sơ đồ thống nhất đa engine (Canvas, PlantUML, Mermaid, Draw.io mxGraph).
    - `mxgraph.py`: Bộ sinh sơ đồ cấu trúc Draw.io XML native (`MX_INV_01`..`20`) và tự động nhúng vào `.docx` / `.pptx`.
    - `code_to_diagram.py`: Bộ quét mã nguồn tĩnh Python AST & SQL/Prisma DDL sinh sơ đồ Class/ERD tự động kèm bộ xuất Vector SVG thuần Python (`export_mxgraph_to_svg`).
    - `editor.py`: Trình biên tập Canvas tương tác 2 chiều.
    - `legacy/`: Các renderer kế thừa (`canvas`, `plantuml`, `mermaid`).

### 2.2. `server/` — Tầng Giao Thức & Cung Cấp Dịch Vụ (Web API, MCP, SSE)

- `app.py`: REST API Server nhẹ nhàng dựa trên `http.server`, phục vụ Web Studio cổng `8000`. Cung cấp endpoints cho conversion, history, mail merge, ocr, diff, math, static assets, trọn bộ 17 công cụ PDF Pro Endpoints, Atomic Office Endpoints, Formula Playbook Endpoints, trọn bộ Enterprise Printing Endpoints (`/api/print/printers`, `/api/print/status`, `/api/print/job`), và System Maintenance Endpoints (`/api/system/maintenance`, `/api/system/vitals`).
- `mcp_server.py`: Model Context Protocol Server (STDIO JSON-RPC 2.0) tương thích với Claude Desktop, Cursor, và Antigravity IDE. Đăng ký **43 tools** chuẩn:
  - Macro Tools (11): `env_doctor`, `doc_convert`, `doc_inspect`, `doc_insert_diagram`, `diagram_render`, `pptx_synthesize`, `docx_apply_preset`, `doc_mail_merge`, `doc_redline_diff`, `doc_ocr_scan`, `pdf_quick_tools`.
  - Atomic Office Tools (8): `xlsx_set_cell`, `xlsx_set_range`, `xlsx_insert_rows`, `xlsx_find_anchor`, `docx_inspect_placeholders`, `docx_replace_text`, `docx_insert_paragraph`, `docx_insert_table`.
  - Excel Power Suite & Diff Tools (6): `xlsx_power_clean`, `xlsx_transform_columns`, `xlsx_pivot_chart`, `xlsx_xlookup_join`, `xlsx_split_by_column`, `xlsx_diff_visual`.
  - Word Mastery Toolkit Tools (4): `docx_polish`, `docx_merge`, `docx_split`, `docx_watermark_hf`.
  - Formula Playbook & Offline AST Evaluator Tools (2): `xlsx_formula_helper`, `xlsx_evaluate_formulas`.
  - Document Security Tool (1): `doc_security_audit`.
  - Enterprise Omni-Pipeline & Reporting Tools (5): `pdf_crypto_sign_and_pdfa`, `xlsx_to_brief_report`, `xlsx_consolidate_workbooks`, `doc_run_pipeline`, `diagram_from_code`.
  - Enterprise Printing Tools (3): `doc_list_printers`, `doc_printer_status`, `doc_print_job`.
  - Universal RAG & Hybrid AI Copilot Tools (2): `universal_rag_search`, `copilot_assist`.
  - System Maintenance Tool (1): `system_maintenance`.
- `job_manager.py`: Server-Sent Events (SSE) background task manager và job queue, cho phép theo dõi tiến trình OCR, mail merge, render và in ấn theo thời gian thực.
- `job_queue.py`: Hàng đợi tác vụ bền bỉ trên SQLite (`PersistentJobQueue`) tuân thủ chốt chặn **G1 (`PRAGMA journal_mode = WAL; PRAGMA busy_timeout = 5000;`)**, cơ chế Retry Exponential Backoff và bộ quét phục hồi sau sự cố (`recover_stale_jobs`).
- `routes/`: Hệ thống định tuyến phân tầng (`Router` Singleton) gồm `system_routes.py`, `print_routes.py`, `doc_routes.py`, `xlsx_routes.py`, `docx_routes.py`, và `copilot_routes.py`.
- `db.py`: Zero-config SQLite database (`history.db`) lưu trữ lịch sử chuyển đổi và tải tệp.

### 2.3. `cli/` — Tầng Giao Diện Dòng Lệnh (CLI Layer)

`cli/ai_tools.py` và facade tại root `ai_tools_cli.py` cung cấp bộ subcommands hoàn chỉnh cho AI Agents và terminal users:

- `doctor`: Kiểm tra sức khỏe môi trường.
- `convert`: Chuyển đổi phổ quát 5 định dạng (PDF, DOCX, MD, PPTX, XLSX).
- `md-toc`: Tạo, xem trước và tự động đồng bộ hóa mục lục tương tác trực tiếp trong tệp Markdown.
- `docx-read`, `docx-write`, `docx-inject`, `docx-patch`, `docx-diff`: Quản lý tài liệu Word.
- `docx-tokens`, `docx-replace`: Atomic Word CLI tools (quét biến & thay thế text bảo tồn OpenXML run).
- `docx-polish`, `docx-merge`, `docx-split`, `docx-watermark`: Word Mastery CLI tools (chuẩn NĐ 30/2020, ghép/tách theo heading ZIP, đóng dấu Watermark + Header/Footer).
- `xlsx-set`, `xlsx-insert-rows`, `xlsx-eval`: Atomic Excel CLI tools (sửa ô hợp nhất, mở rộng dòng kèm dịch công thức, và giải công thức AST ngoại tuyến nạp thẻ `<v>`).
- `xlsx-power-clean`, `xlsx-transform-cols`, `xlsx-pivot-chart`, `xlsx-xlookup-join`, `xlsx-split-col`, `xlsx-diff`, `xlsx-track`: Excel Power Suite & Work Tracker CLI tools (`cli/office_power_commands.py`).
- `xlsx-formulas`: Cẩm nang tra cứu và xuất Excel Formula Playbook 57+ công thức.
- `doc-security`: Kiểm tra và tẩy rửa an ninh tài liệu, formula injection, OLE PE binary, và PII (Cloudflare Harness).
- `docx-mail-merge`: Trộn thư hàng loạt Word + Excel/JSON (hỗ trợ filter `:vn_words`, `:date_vn`, `:vnd`).
- `docx-apply-preset`, `docx-sanitize-vietnamese`: Chuẩn hóa thể thức hành chính & chính tả.
- `pptx-build`, `pptx-inspect`, `pptx-preview`, `pptx-insert-diagram`, `docx-to-pptx`, `pdf-to-pptx`: Quản lý thuyết trình PowerPoint.
- `pdf-tools`: Bộ 17 công cụ PDF Pro Studio (`merge`, `split`, `rotate`, `clean`, `watermark`, `compress`, `redact`, `protect`, `normalize-a4`, `giap-lai`, `extract-tables`, `images-to-pdf`, `export-images-zip`, `sign`, `flatten`, `reorganize`, `crypto-sign`, `pdfa`).
- `ocr-scan`: Bóc tách tài liệu scan và loại bỏ dấu mộc đỏ.
- `diagram-render`, `plantuml-render`, `mermaid-render`, `spec-render`, `diagram-editor`: Sơ đồ kỹ thuật.
- `rag-index`, `rag-search`, `copilot-chat`: Lập chỉ mục và truy xuất ngữ nghĩa đa định dạng qua Universal SQLite FTS5 RAG và hỏi đáp với Hybrid AI Copilot kèm Thẻ Hành Động.
- `print-list`, `print-status`, `print-doc`: Liệt kê máy in hệ thống, tra cứu trạng thái phần cứng (hết giấy, kẹt giấy, offline, số lệnh chờ) và thực hiện in ngầm với bộ tối ưu hóa in ấn (Dải trang Page Range, Khổ giấy A4/Letter/A3, Căn lề Fit/Actual, Standard Duplex, Vector Grayscale `/DeviceGray`, N-Up, Booklet Creep Shingling, DEVMODE Tray/Duplex Guard G2, Margin Guard).
- `system-maintenance`: Bộ công cụ bảo trì hệ thống 7 trụ cột (dọn lock files ~$*, kết liễu tiến trình mồ côi, vacuum SQLite, dọn cache, chẩn đoán Windows MPO).
- `qa-audit`: Master QA diagnostic and repair pipeline 13 Cổng (`tools/unified_qa_diagnostic.py`).

### 2.4. `ui/` & `src/` — Tầng Giao Diện Người Dùng (Desktop GUI & Web Studio)

- `ui/`: Desktop GUI Tkinter độc lập (`main.py`) và Native WebView App Mode (`ui/webview_launcher.py`) hỗ trợ chạy offline không cần trình duyệt.
- `src/`: Web Studio trên nền **React 18 + TypeScript + Tailwind CSS v3 + `lucide-react` + Vite 6** (Light-First Enterprise Office + Dark Mode, thuần Local `127.0.0.1:8000`):
  - `src/App.tsx`, `src/main.tsx`, `src/styles/index.css`: Root Shell (Collapsible Left Sidebar, HeaderBar kèm `Ctrl+K` Command Palette, Session File Tray, Inline Document Preview, AI Copilot Dock khóa Local RAG mặc định, và 1-Click Xóa vết phiên `/api/session/purge`).
  - `src/context/`, `src/services/api.ts`, `src/types/`, `src/config/`: `ThemeContext`, `ToastContext`, `SessionFilesContext`, `WorkspaceContext`, typed API client & navigation/diagram registries.
  - `src/components/`: UI primitives (`ui/`), layout (`SplitWorkspace`, `HeaderBar`, `LeftSidebar`, `SessionFileTray`, `CommandPalette`), `InlinePreviewPane` (PDF/DOCX/XLSX/Ảnh), và global modals (`CopilotSidebar`, `PrintStudioModal`, `HistoryDrawer`, `SystemDiagnosticsModal`, `QuickGuideModal`).
  - `src/features/`: Các phân hệ Pure Office: `home/`, `word/` (5 công cụ), `pdf/` (PDF Pro 18 nghiệp vụ + Ký vẽ tay + Sắp xếp trang 120 DPI), `excel/` (Power Suite, Reporting, Tracker, 57 Công thức, Template & Print), `convert/` (Chuyển đổi đa chiều, OCR Tiếng Việt, QA 13 cổng & Deep CDR), và `visual/` (Slide 16:9 & Diagram Studio).

### 2.5. `tools/` — Tầng Tiện Ích QA & Giao Nộp (Diagnostics & Deliverables)

- `tools/unified_qa_diagnostic.py`: Động cơ chẩn đoán toàn diện 13 Cổng cho DOCX, XLSX, PPTX, Draw.io, và ảnh scan với khả năng tự sửa lỗi (Safe Non-destructive Repair).
- `tools/format_diff_excel.py`: Tiện ích so khớp định dạng 12 cổng bảng tính Excel.
- `tools/bump_version.py` & `tools/sync_core_docs.py`: Tự động đồng bộ 10 điểm neo phiên bản (`ERR_VER_001`) và kiểm định đối xứng ngữ nghĩa bộ tài liệu lõi (`ERR_DOC_SYNC_001`).
- `tools/gen_*.py`: Các bộ sinh sơ đồ kỹ thuật chuẩn (`gen_canvas_erd.py`, `gen_hub_spoke_flow.py`, `gen_master_erd.py`, `gen_usecase_diagram.py`, `gen_illustration_assets.py`).
- `tools/deliverables/`: Lưu trữ riêng các scripts xuất báo cáo/đồ án cá nhân (được loại trừ trong `.gitignore` để giữ kho mã nguồn chính luôn gọn sạch).

---

## 3. Tiêu Chuẩn Clean-Room MIT & Chế Độ Thuần Local Bảo Mật (`127.0.0.1`)

Dự án được thiết kế theo chuẩn Clean-Room và vận hành **thuần máy trạm cục bộ (Pure Local-Only)** để bảo mật dữ liệu văn phòng tối đa:

1. **Chuỗi Fallback Zero-COM & Thực Thi Cục Bộ**:
   - Trên Windows: Ưu tiên Win32 COM nếu có MS Office và không bị khóa tiến trình.
   - Dự phòng đa nền tảng: Tự động chuyển tiếp sang LibreOffice Headless (`soffice --headless --convert-to pdf`) hoặc thư viện Python thuần (`PyMuPDF`, `openpyxl`, `python-docx`, `python-pptx`).
2. **Bảo Mật Thuần Local (`127.0.0.1:8000`)**:
   - Server REST API khóa mặc định vào loopback `127.0.0.1:8000`, tường lửa Origin chặn mọi yêu cầu cross-origin ngoài `localhost`/`127.0.0.1`, AI Copilot khóa mặc định Local/Offline RAG và hỗ trợ dọn sạch dấu vết phiên qua `/api/session/purge`.

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
   Tất cả bài kiểm thử tự động trong `tests/` phải luôn luôn PASS 100% trước khi phát hành.
3. **Giữ Vững CLI Contract**:
   Tất cả cú pháp lệnh trong `AGENTS.md` và `pyproject.toml` tiếp tục hoạt động chính xác không thay đổi.
