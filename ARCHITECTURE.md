# Architecture Guide — Antigravity Universal Office Studio

> **Phiên bản**: 4.2.0 (Target Layered Modular Architecture v4.2)  
> **Tiêu chuẩn**: Zero-Regression, 100% Backward Compatibility, Strict Layer Isolation, Clean-Room MIT Compliance  
> **Test Suite**: 220 Automated Tests (100% Pass Rate)

---

## 1. Tổng Quan Kiến Trúc Phân Tầng (5-Layer Modular Architecture)

Hệ thống được tổ chức theo mô hình phân tầng chặt chẽ (Clean Architecture & Decoupled Modular Layers):

```
┌────────────────────────────────────────────────────────────────────────────────────────┐
│                              1. Presentation & Client Layer                            │
│   Web Studio (src/ - Vite ES6 + Vanilla Glassmorphic) │ Desktop GUI (ui/ - Tkinter)    │
└───────────────────────────────────────────┬────────────────────────────────────────────┘
                                            │
                                            ▼
┌────────────────────────────────────────────────────────────────────────────────────────┐
│                              2. Interface & Protocol Layer                             │
│   Web REST API (server/app.py :8000)      │ Model Context Protocol (server/mcp_server) │
│   SSE Task Manager (server/job_manager)   │ AI Unified CLI (cli/ai_tools.py)           │
└───────────────────────────────────────────┬────────────────────────────────────────────┘
                                            │
                                            ▼
┌────────────────────────────────────────────────────────────────────────────────────────┐
│                            3. Core Document & Automation Engine                        │
│  core/docx/     │ core/pptx/     │ core/xlsx/     │ core/converter/ │ core/pdf/        │
│  core/diagrams/ │ core/env_doctor│                                                     │
└───────────────────────────────────────────┬────────────────────────────────────────────┘
                                            │
                                            ▼
┌────────────────────────────────────────────────────────────────────────────────────────┐
│                          4. Storage, Cache & Diagnostic Layer                          │
│   SQLite (history.db / db.py) │ In-Memory SSE Queue │ Unified QA Diagnostic Engine     │
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
    - `docx_to_pptx.py`: Bộ chuyển đổi thông minh từ Word sang slide thuyết trình 16:9, tự động nhận diện bố cục doanh nghiệp và co giãn đối xứng các khối nội dung.
    - `math.py`: Native DrawingML Math (`baseline="-25000"`, `baseline="30000"`, biến in nghiêng).
    - `charts.py`: Biểu đồ OpenXML Native (`<c:chart>`) nhúng gói bảng tính Excel ẩn cho phép "Edit Data" trực tiếp trên PowerPoint (`PPTX_INV_14`).
    - Business Layouts: Hỗ trợ 4 mẫu slide doanh nghiệp (KPI Summary Cards, Timeline Horizontal Roadmap, SWOT 2x2 Matrix, Process Chevron Steps).
  - `core/xlsx/`:
    - `reader.py` / `writer.py`: Đọc, mutate bảng tính Excel in-place, clone style prototype row (`ERR_XLSX_001`), dịch chuyển công thức động (`ERR_XLSX_002`), thiết lập trang in Print-Ready (`ERR_XLSX_008`).
    - `validator.py`: Kiểm định chất lượng 12 Cổng (Invariant E6).
    - `format_diff.py`: Bóc tách và so sánh định dạng giữa các sheet.
  - `core/converter/`:
    - `engine.py`: Dispatcher ma trận 4 chiều (PDF ↔ DOCX ↔ MD ↔ PPTX) và Decoupled Markdown parser.
    - `ocr_pipeline.py`: Pipeline bóc tách PDF/ảnh scan tiếng Việt (`vie+eng`), bộ lọc mộc đỏ HSV (`ERR_CONV_005`), nắn thẳng ảnh scan tự động (`deskew`).
    - `pdf_renderer.py`: Bộ chuyển đổi PDF phổ quát (`UniversalPdfRenderer`) với chuỗi ưu tiên: Win32 COM ➔ LibreOffice Headless ➔ PyMuPDF.
    - `pdf_toolkit.py`: Phân hệ PDF Pro Studio & iLovePDF Killer Features (True Redaction `ERR_PDF_001`, AES-256 Encryption, Dấu giáp lai đa trang overlap 8% `ERR_PDF_002`, Chuẩn hóa A4 vector `ERR_PDF_003`, Bóc bảng sang Excel `ERR_XLSX_007`, Tổ chức trang trực quan kèm chèn trang trắng A4, Chuyển đổi 2 chiều Ảnh ↔ PDF, Ký tên điện tử & đóng dấu bù góc xoay `page.derotation_matrix`, Làm phẳng 2 chế độ Vector Bake & Strict Raster kèm dọn sạch metadata ẩn `doc.scrub`).
  - `core/env_doctor/`:
    - `run_doctor.py`: Tự động quét kiểm tra 12 thành phần môi trường (Python, LibreOffice, Tesseract, Chromium, PlantUML, font chữ) và xuất báo cáo JSON / Terminal kèm mã lệnh khắc phục 1-click.
  - `core/diagrams/`:
    - `bridge.py`: Điều phối sơ đồ thống nhất.
    - `mxgraph.py`: Bộ sinh sơ đồ cấu trúc Draw.io XML native (`MX_INV_01`..`20`).
    - `editor.py`: Trình biên tập Canvas tương tác 2 chiều.
    - `legacy/`: Các renderer kế thừa (`canvas`, `plantuml`, `mermaid`).

### 2.2. `server/` — Tầng Giao Thức & Cung Cấp Dịch Vụ (Web API, MCP, SSE)
- `app.py`: REST API Server nhẹ nhàng dựa trên `http.server`, phục vụ Web Studio cổng `8000`. Cung cấp endpoints cho conversion, history, mail merge, ocr, diff, math, và static assets. Bổ sung trọn bộ PDF Pro Endpoints: `POST /api/pdf/pages/preview`, `POST /api/pdf/tools`, `POST /api/pdf/images-to-pdf`, `POST /api/pdf/export-images-zip`, `POST /api/pdf/sign`, `POST /api/pdf/sanitize`.
- `mcp_server.py`: Model Context Protocol Server (STDIO JSON-RPC 2.0) tương thích với Claude Desktop, Cursor, và Antigravity IDE. Đăng ký 11 tools chuẩn:
  - `env_doctor`, `doc_convert`, `doc_inspect`, `doc_insert_diagram`, `docx_patch`, `docx_apply_preset`, `docx_sanitize_vietnamese`, `pptx_synthesize`, `diagram_render`, `doc_mail_merge`, `doc_redline_diff`, `doc_ocr_scan`, `pdf_quick_tools`.
- `job_manager.py`: Server-Sent Events (SSE) background task manager và job queue, cho phép theo dõi tiến trình OCR, mail merge, và render theo thời gian thực.
- `db.py`: Zero-config SQLite database (`history.db`) lưu trữ lịch sử chuyển đổi và tải tệp.

### 2.3. `cli/` — Tầng Giao Diện Dòng Lệnh (CLI Layer)
`cli/ai_tools.py` và facade tại root `ai_tools_cli.py` cung cấp 23 subcommands hoàn chỉnh cho AI Agents và terminal users:
- `doctor`: Kiểm tra sức khỏe môi trường.
- `docx-read`, `docx-write`, `docx-inject`, `docx-patch`, `docx-diff`: Quản lý tài liệu Word.
- `docx-mail-merge`: Trộn thư hàng loạt Word + Excel/JSON.
- `docx-apply-preset`, `docx-sanitize-vietnamese`: Chuẩn hóa thể thức hành chính & chính tả.
- `pptx-build`, `pptx-inspect`, `pptx-preview`, `pptx-insert-diagram`, `docx-to-pptx`, `pdf-to-pptx`: Quản lý thuyết trình PowerPoint.
- `pdf-tools`: Bộ 16 công cụ PDF Pro Studio (`merge`, `split`, `rotate`, `clean`, `watermark`, `compress`, `redact`, `protect`, `normalize-a4`, `giap-lai`, `extract-tables`, `images-to-pdf`, `export-images-zip`, `sign`, `flatten`, `reorganize`).
- `ocr-scan`: Bóc tách tài liệu scan và loại bỏ dấu mộc đỏ.
- `diagram-render`, `plantuml-render`, `mermaid-render`, `spec-render`, `diagram-editor`: Sơ đồ kỹ thuật.
- `qa-audit`: Master QA diagnostic and repair pipeline (`tools/unified_qa_diagnostic.py`).

### 2.4. `ui/` & `src/` — Tầng Giao Diện Người Dùng (Desktop GUI & Web Studio)
- `ui/`: Desktop GUI Tkinter độc lập (`main.py`) hỗ trợ chạy offline không cần trình duyệt.
- `src/`: Web Studio trên nền Vite ES6 modules và Vanilla CSS Glassmorphism:
  - Tách rời theo module: `document_hub.js`, `presentation_studio.js`, `diagram_studio.js`, `mail_merge_studio.js`, `math_studio.js`, `history_drawer.js`.

### 2.5. `tools/` — Tầng Tiện Ích QA & Giao Nộp (Diagnostics & Deliverables)
- `tools/unified_qa_diagnostic.py`: Động cơ chẩn đoán toàn diện DOCX, XLSX, PPTX, Draw.io, và ảnh scan với khả năng tự sửa lỗi (Safe Non-destructive Repair).
- `tools/format_diff_excel.py`: Tiện ích so khớp định dạng 12 cổng bảng tính Excel.
- `tools/deliverables/`: Scripts xuất báo cáo chuyên biệt (Report3, Report5, Luận văn).

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

## 5. Bất Biến Tương Thích Ngược (Backward-Compatibility Guarantees)

1. **Thin Facades Tại Thư Mục Gốc**:
   Mọi module cũ tại root (`docx_reader.py`, `docx_writer.py`, `pptx_writer.py`, `web_api_server.py`, `mcp_server.py`, `ai_tools_cli.py`, v.v.) được duy trì dưới dạng **Thin Facades**, re-export toàn bộ symbol từ `core/`, `server/`, `ui/`.
2. **100% Pass Rate Toàn Bộ Test Suite**:
   Tất cả 193 bài kiểm thử tự động trong `tests/` phải luôn luôn PASS 100% trước khi phát hành.
3. **Giữ Vững CLI Contract**:
   Tất cả cú pháp lệnh trong `AGENTS.md` tiếp tục hoạt động chính xác không thay đổi.
