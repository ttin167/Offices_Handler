# Architecture Guide — Antigravity Universal Office Studio

> **Phiên bản**: 3.2.0 (Layered Modular Architecture)  
> **Tiêu chuẩn**: Zero-Regression, 100% Backward Compatibility, Strict Layer Isolation

---

## 1. Tổng Quan Kiến Trúc Phân Tầng (Layered Architecture)

Hệ thống được tổ chức theo mô hình phân tầng chặt chẽ (Clean Architecture / Decoupled Layers):

```
┌──────────────────────────────────────────────────────────────────────────┐
│                         Presentation & Client Layer                      │
│   Web Studio (src/ - Vite + CSS)   │  Desktop GUI (ui/ - Tkinter)       │
└────────────────────────────────────┴─────────────────────────────────────┘
                                      │
                                      ▼
┌──────────────────────────────────────────────────────────────────────────┐
│                           Interface & CLI Layer                          │
│   Web API Server (server/ - :8000) │  AI Tools CLI (cli/ - ai_tools_cli) │
└────────────────────────────────────┴─────────────────────────────────────┘
                                      │
                                      ▼
┌──────────────────────────────────────────────────────────────────────────┐
│                         Core Document Engine Layer                       │
│  core/docx/  │  core/pptx/  │  core/xlsx/  │  core/converter/  │  core/diagrams/
└──────────────────────────────────────────────────────────────────────────┘
                                      │
                                      ▼
┌──────────────────────────────────────────────────────────────────────────┐
│                       Storage & External Tooling Layer                   │
│   SQLite (history.db)  │  OpenXML / DrawingML  │  plantuml.jar / Chromium│
└──────────────────────────────────────────────────────────────────────────┘
```

---

## 2. Trách Nhiệm Từng Tầng (Layer Responsibilities)

### 2.1. `core/` — Tầng Lõi Nghiệp Vụ Tài Liệu (Core Engines)
Tầng này chứa 100% logic thuật toán xử lý dữ liệu và tạo lập tài liệu văn phòng / đồ họa.
- **Quy tắc phụ thuộc**: Tuyệt đối không import từ `server/`, `ui/`, hoặc `cli/`. Chỉ nhận tham số thuần túy (Python dict, dataclass, file path) và trả về kết quả.
- **Phân hệ**:
  - `core/docx/`: Đọc, sinh, vá mẫu OpenXML Word. Bảo toàn thuộc tính `<w:cantSplit/>`, `<w:tblHeader/>`, `<w:vAlign w:val="center"/>`.
  - `core/pptx/`: Đọc, biên dịch slide Widescreen 16:9, bộ theme 4K, Native DrawingML device mockups (macOS Browser, Smartphone iOS/Android).
  - `core/xlsx/`: Đọc, mutate bảng tính Excel in-place, kiểm định 12 Cổng Chất Lượng (Invariant E6).
  - `core/converter/`: Dispatcher ma trận 4 chiều (PDF ↔ DOCX ↔ MD ↔ PPTX) và Decoupled Markdown parser.
  - `core/diagrams/`: Động cơ sơ đồ kiến trúc Draw.io XML (`mxgraph.py`), interactive canvas editor (`editor.py`), và các động cơ legacy (`legacy/`).

### 2.2. `server/` — Tầng Web API Bridge
Cung cấp REST API nội bộ không phụ thuộc framework nặng (sử dụng thư viện chuẩn `http.server`), phục vụ Web Studio trên cổng `8000`.
- Chịu trách nhiệm: định tuyến request, parse multipart upload, quản lý tiến trình nền (spawning Chromium, PlantUML), và lưu trữ lịch sử qua SQLite (`db.py`).

### 2.3. `ui/` — Tầng Desktop GUI
Cung cấp giao diện đồ họa độc lập trên nền Python Tkinter cho người dùng không muốn chạy trình duyệt.

### 2.4. `cli/` — Tầng Giao Diện Dòng Lệnh
Cung cấp công cụ dòng lệnh cho AI Agents (`ai_tools.py`) và lập trình viên terminal:
- Áp dụng cơ chế **Dual-Fallback Import**:
  ```python
  try:
      from core.docx.reader import read_docx_to_json
  except (ImportError, ValueError):
      from docx_reader import read_docx_to_json
  ```
- Cho phép chạy mượt mà ở cả 3 chế độ:
  1. `python ai_tools_cli.py <cmd>`
  2. `python -m ai_tools_cli <cmd>`
  3. Console script entrypoint: `doc-handler <cmd>`

### 2.5. `src/` — Tầng Giao Diện Web Studio
Xây dựng trên nền Vite ES6 modules và Vanilla CSS:
- Tách rời theo module tính năng: `api.js`, `state.js`, `presentation_studio.js`, `document_hub.js`, `diagram_studio.js`, `inspector_modal.js`.
- Không sử dụng thư viện UI nặng nề (giữ vững nguyên tắc Vanilla Glassmorphic).

### 2.6. `tools/` — Tiện Ích QA & Giao Nộp
- `tools/unified_qa_diagnostic.py`, `tools/format_diff_excel.py`: Tiện ích QA dùng chung cho CI/CD.
- `tools/deliverables/`: Nơi lưu trữ các script sinh báo cáo cụ thể (Report3, Report5, Luận văn) để giữ sạch không gian làm việc.

---

## 3. Bất Biến Tương Thích Ngược (Backward-Compatibility Invariants)

1. **Facade Stubs Tại Thư Mục Gốc**:
   Mọi module cũ tại root (`docx_reader.py`, `docx_writer.py`, `pptx_writer.py`, `web_api_server.py`, v.v.) được duy trì dưới dạng **Thin Facades**, re-export toàn bộ các symbol từ `core/`, `server/`, `ui/`.
2. **Không Làm Gãy Test Suite**:
   Toàn bộ 108 unit tests trong `tests/` phải luôn luôn PASS 100% trong suốt quá trình và sau khi hoàn thành tái cấu trúc.
3. **Giữ Vững Các Lệnh Trong AGENTS.md**:
   Tất cả các cú pháp lệnh tài liệu hóa trong `AGENTS.md` (`python ai_tools_cli.py ...`, `python main.py`) tiếp tục hoạt động chính xác.
