# GEMINI.md — Antigravity Agent Critical Rules & Fast-Dispatch Manual

> **Auto-loaded every Antigravity session.**
> Master Catalog: `AGENTS.md` | Architecture: `ARCHITECTURE.md` | Session Task State: `.agent_scratchpad.md`

---

## 1. CHECKPOINT Recovery Protocol (BẮT BUỘC)

Khi nhận tín hiệu `{{ CHECKPOINT N }}` hoặc bắt đầu một phiên làm việc mới:
1. **Đọc file này (`GEMINI.md`)** trước tiên — xác nhận toàn bộ invariants kỹ thuật đã nạp vào ngữ cảnh.
2. **Đọc `.agent_scratchpad.md`** — nắm bắt `Current Goal`, `Progress / Completed Steps`, và `Remaining TODOs`.
3. Không tự ý chạy lệnh đoán mò — tham khảo bảng **Fast-Dispatch Matrix** bên dưới để mở thẳng file nguồn cần thao tác.

---

## 2. 🗺️ Fast-Dispatch Direct File Matrix (AI Direct Action)

> [!IMPORTANT]
> **CRITICAL DIRECTIVE**:
> Nghiêm cấm chạy các lệnh quét thư mục đệ quy (`Get-ChildItem -Recurse`, `find`, `grep` bừa bãi). Tra cứu ma trận bên dưới và mở trực tiếp tệp mã nguồn:

| Nhiệm vụ / Domain | Tệp Mã Nguồn Cốt Lõi | Hàm / Lớp Điểm Neo |
|---|---|---|
| **System Diagnostics & Health** | `core/env_doctor.py` | `run_doctor()` |
| **Universal Lossless Converter** | `core/converter/engine.py` | `convert_universal()` |
| **PDF Toolkit & PAdES Sign** | `core/converter/pdf_toolkit.py` | `PDFQuickToolkit`, `crypto_sign_pdf` |
| **Document AST Parser (DOCX/PPTX)** | `core/docx/reader.py`, `core/pptx/reader.py` | `read_docx_to_json()`, `read_pptx()` |
| **Legal Track Changes Redline Diff** | `core/docx/diff_engine.py` | `compare_docx_documents()` |
| **Mail Merge & Templating** | `core/docx/mail_merge.py`, `core/docx/writer.py` | `inspect_template_tokens()`, `patch_docx_template()` |
| **13-Gate Unified QA Diagnostic** | `core/qa/unified_audit.py` | `run_unified_qa()` |
| **Vietnamese OCR & Scan Recovery** | `core/converter/ocr_pipeline.py` | `VietnameseOcrPipeline` |
| **Multi-Engine Diagram Render** | `core/diagrams/bridge.py` | `render_diagram_file_to_png()` |
| **PowerPoint 16:9 Deck Builder** | `core/pptx/docx_to_pptx.py`, `core/pptx/engine.py` | `convert_docx_to_pptx()` |
| **Excel In-Place Mutation Engine** | `core/xlsx/writer.py` | `UniversalExcelEngine`, `mutate_template_excel()` |
| **Offline AST Formula Evaluator** | `core/xlsx/formula_evaluator.py` | `ASTFormulaEvaluator.populate_cached_values()` |
| **Cloudflare Security & Deep CDR** | `core/security/document_auditor.py` | `SecurityAuditor`, `OfficeSanitizer` |
| **Declarative Workflow Pipelines** | `core/pipeline/engine.py` | `WorkflowPipelineEngine` |
| **AST Code-to-Diagram Engine** | `core/diagrams/code_to_diagram.py` | `generate_diagram_from_code()` |
| **Excel Power Suite (Clean/Pivot/XLOOKUP)** | `core/xlsx/power_tools.py` | `clean_and_standardize_workbook()`, `build_pivot_summary_and_chart()`, `xlookup_join_workbooks()` |
| **Office Work & Daily Task Tracker** | `core/xlsx/tracker.py` | `OfficeWorkTracker` |
| **Excel 57-Formula Playbook** | `core/xlsx/formula_playbook.py` | `FORMULA_CATALOG`, `generate_formula_playbook_xlsx()` |
| **Word Mastery Toolkit (NĐ30/Merge/Split)** | `core/docx/toolkit.py` | `polish_docx_document()`, `merge_docx_files()`, `split_docx_by_heading()`, `apply_docx_watermark_and_hf()` |
| **Vietnamese Amount in Words (`vn_words`)** | `core/docx/design_system.py` | `number_to_vietnamese_words()`, `format_token_value()` |
| **Enterprise Silent Print Spooler** | `core/printing/printer_manager.py` | `PrinterManager.print_document()` |
| **Universal FTS5 RAG Vault** | `core/copilot/universal_rag.py` | `UniversalRAGEngine` |
| **Hybrid AI Copilot Assistant** | `core/copilot/copilot_provider.py` | `CopilotOrchestrator` |
| **7-Pillars Periodic Maintenance** | `core/system_maintenance.py` | `run_periodic_maintenance()` |
| **Model Context Protocol (43 Tools)** | `server/mcp_server.py` | `TOOL_DEFINITIONS`, `handle_rpc_request()` |
| **Modular REST API Routes** | `server/routes/` | `xlsx_routes.py`, `docx_routes.py`, `doc_routes.py`, `print_routes.py`, `copilot_routes.py` |
| **Unified Terminal CLI Dispatcher** | `ai_tools_cli.py`, `cli/` | `office_power_commands.py`, `v46_commands.py`, `print_commands.py` |

---

## 3. 🛡️ Bảng Tra Cứu Invariants Kỹ Thuật Bắt Buộc

### A. Spreadsheet & Excel Invariants (E1–E14)

| # | Invariant | Quy định cốt lõi |
|---|---|---|
| E1 | **Template-Driven Token Extraction** | 100% tokens (font, fill, border, alignment) lấy từ reference sheet (`Example`/`Template`). Không đoán mò. |
| E2 | **Grill-Before-Deviate** | Đề xuất cải tiến UX lệch template -> Dừng lại, hỏi user qua `/grill-me` trước khi áp dụng. |
| E3 | **Live KPI Formulas** | Row 7 và Summary cells phải dùng công thức động (`=COUNTIF`, `=SUM`, `='Sheet'!Cell`), không hardcode số. |
| E4 | **Era-Appropriate Data** | Dữ liệu mock/test phải khớp niên đại dự án (2026+), loại bỏ placeholder cũ 2000-2009. |
| E5 | **Sibling Symmetry** | Toàn bộ function sheets phải đồng nhất 100% về kích thước cột, dòng, freeze panes, và styles. |
| E6 | **Automated 12-Gate Diff** | Bắt buộc chạy script format diff 12 cổng trước khi bàn giao. Nghiêm cấm `exit code 0 = Done`. |
| E7 | **Unified 3-Column Box (B-C-D)** | Khởi tạo đầy đủ khối B-C-D cho mọi hàng dữ liệu: B (`left=thin`), C (no vertical), D (`right=thin`), fill white. |
| E8 | **Strict Group Hierarchy** | Tên nhóm (e.g. `Precondition`) chỉ xuất hiện ở Col B dòng đầu tiên. Các dòng con Col B để trống, text ở Col D. |
| E9 | **Cross-Zone Token Isolation** | Trích xuất token theo từng vùng (Header, Condition, Confirm, Result footer). Không dùng chung style giữa các vùng. |
| E10 | **Multi-Tier Hierarchy** | Giữ nguyên 3 cấp phân cấp: Level 1 (Precondition), Level 2 (Input Parameters/ParamName), Level 3 (Value). |
| E11 | **Untrusted Input Style Isolation** | File input chỉ dùng lấy raw value. Tuyệt đối KHÔNG copy font size/style từ input (tránh lỗi Arial 8.5pt). |
| E12 | **Dynamic Chart Re-Anchoring** | OpenPyXL không tự relink chart -> Bắt buộc script duyệt `chart.series` cập nhật formula `$F$34:$H$34` và anchor row. |
| E13 | **UX Dynamic Text Auto-Scaling** | `Row Height = max(min_h, total_lines * line_h)` + `wrap_text=True` cho các ô text dài, chống tràn/cắt chữ. |
| E14 | **Summary & Subtotal Parity** | Subtotal: Navy endpoints (A/B & Total), nền trắng giữa (C..H). Đủ 5 dòng KPI (Coverage, Success, Normal, Abnormal, Boundary). |

### B. Enterprise Document & Spreadsheet QA Invariants

| Mã Lỗi | Tên Lỗi | Nguyên Tắc Bắt Buộc |
|---|---|---|
| `ERR_XLSX_001` | **Prototype Row Style Cloning** | Luôn clone 100% style từ dòng dữ liệu đại diện trong template sang dòng mới. |
| `ERR_XLSX_002` | **Dynamic Formula Shifting** | Bóc tách dải ô tham chiếu bằng Regex và tịnh tiến động khi phình to dòng. |
| `ERR_XLSX_003` | **Semantic Anchor Discovery** | Dò tìm Header, Data Start và Summary bằng từ khóa; dữ liệu chỉ chèn vào giữa. |
| `ERR_XLSX_004` | **Safe Merged-Cell Handling** | Chỉ gán giá trị vào Top-Left; đồng bộ style toàn dải merged chống rách viền. |
| `ERR_XLSX_005` | **Freeze Panes & Dynamic Height** | Freeze Panes tại giao điểm Header+1; Chiều cao tính bằng `max(min_h, lines * line_h)`. |
| `ERR_XLSX_006` | **DrawingML Preservation** | Luôn load trực tiếp template gốc; cấm tạo `Workbook()` rỗng làm mất shapes/logo. |
| `ERR_XLSX_007` | **Identifier Number Format** | Cột mã định danh (ID, Code) bắt buộc gán `number_format = '@'` và kiểu string. |
| `ERR_XLSX_015` | **Dual-Workbook Formula Fallback** | Đọc `data_only=True` trả về `None` -> fallback đọc công thức từ `data_only=False` và tự tính `SUM`, `AVERAGE`, `COUNT`, `MIN`, `MAX`. |
| `ERR_DOCX_001` | **The Last Paragraph Rule** | Mọi cell bảng (`<w:tc>`) bắt buộc phải kết thúc bằng tối thiểu một thẻ `<w:p>`. |
| `ERR_DOCX_002` | **Run Text Overwrite** | Thao tác nội dung qua `cell.paragraphs[0].runs`, không gán `cell.text = "..."`. |
| `ERR_DOCX_003` | **Multi-Page Table Flags** | Bảng nhiều trang bắt buộc có `<w:cantSplit/>` và `<w:tblHeader/>`. |
| `ERR_DOCX_005` | **Printable Margin Overflow** | Khóa tỉ lệ ảnh; chiều rộng hình ảnh $\le$ `page_width - left_margin - right_margin` ($\le 14\text{cm}$). |
| `ERR_VER_001` | **Version Desync Invariant** | Luôn đồng bộ 100% version trên 10 điểm neo qua `tools/bump_version.py` + `npm run build`. |
| `ERR_DOC_SYNC_001` | **Feature & Documentation Invariant** | Hoàn tất tính năng bắt buộc cập nhật `CHANGELOG.md`, `ARCHITECTURE.md` và Core Docs Triad (`tools/sync_core_docs.py --build`). |

### C. Draw.io & mxGraphModel Critical Invariants (MX_INV_01–20)

| Mã Lỗi | Tên Bất Biến | Nguyên Tắc Bắt Buộc |
|---|---|---|
| `MX_INV_01` | **Pure Native Hierarchy** | Cấm `<UserObject mermaidData/plantUmlData>`. Mọi table và edge phải là `<mxCell>` trực tiếp dưới `parent="1"`. |
| `MX_INV_02` | **Minimalist & Well-Formed XML** | Xóa metadata rác (`mermaidBaseStyle`, `mermaidId`, `alternateBounds`). File < 300KB, đủ 100% thẻ đóng. |
| `MX_INV_03` | **Container-Level Docking** | Edge chỉ nối từ `table_X` sang `table_Y`, nghiêm cấm nối vào cell con (`tableRow`/`partialRectangle`). |
| `MX_INV_04` | **Orthogonal Perimeter Routing** | `edgeStyle=orthogonalEdgeStyle;` + cổng viền chuẩn (`exitX, exitY, entryX, entryY` là `0.0`, `0.5`, `1.0`). |
| `MX_INV_05` | **Dynamic Geometry Scaling** | Chiều cao $H = 43 \times (N_{\text{fields}} + 1)$, hành lang giao thông an toàn giữa các bảng $\ge 60$px. |
| `MX_INV_07` | **Monochrome Academic Line-Art** | Mọi sơ đồ học thuật & kỹ thuật ưu tiên Monochrome: `#ffffff` fill, `#000000` text/viền/mũi tên. |
| `MX_INV_08` | **Collision-Free Labels & Masks** | Text nhãn trên line bắt buộc có `labelBackgroundColor=#ffffff;` chống đè chữ; hành lang chạy line $\ge 40$px. |
| `MX_INV_11` | **Fractional Perimeter Port Anchoring** | Nối Hub sang vệ tinh dùng toạ độ vi phân chu vi (`exitY = 0.05..1.0`) khớp $y_{\text{center}}$ đích để tạo đường ngang phẳng 100% (Zero-Zigzag). |
| `MX_INV_15` | **Dynamic Edge Label Width** | Cấm `\n`/`<br>` trong label; bắt buộc `labelWidth=<W>;html=1;whiteSpace=wrap;labelBackgroundColor=#FFFFFF;`. |
| `MX_INV_20` | **Snug Mask Bounding** | CẤM gán cứng `labelWidth` lớn cho nhãn ngắn. `labelWidth` phải tính động để mask trắng ôm khít chữ. |

### D. PowerPoint Presentation Invariants (PPTX_INV_01–13)

| Mã Lỗi | Tên Bất Biến | Nguyên Tắc Bắt Buộc |
|---|---|---|
| `PPTX_INV_01` | **Widescreen Master** | Mọi slide mặc định tỷ lệ 16:9 (`13.333 x 7.5 in` / `1280 x 720 pt`). |
| `PPTX_INV_02` | **Aspect-Ratio Lock** | Sơ đồ kỹ thuật và hình ảnh khóa tỉ lệ khung hình ($W/H$); `Width <= 12.0 in`, `Height <= 5.5 in`. |
| `PPTX_INV_03` | **Template Style Inheritance** | Kế thừa đúng font chữ, kích cỡ header, màu sắc từ Slide Master khi truyền `--template`. |
| `PPTX_INV_04` | **Zero Text Overflow** | Text box thiết lập `word_wrap = True` và tự điều chỉnh font size chống tràn chữ. |
| `PPTX_INV_07` | **Windows File Lock Guard** | Slide đang mở trong PowerPoint -> tự động lưu ra file tạm `_updated.pptx` kèm cảnh báo, không crash. |
| `PPTX_INV_09` | **Run-Level Formatting Supremacy** | CẤM gán `b="1"` hay `<a:solidFill>` vào `<a:defRPr>`. Mọi bold/color phải đặt ở `<a:rPr>`. |
| `PPTX_INV_10` | **Default Pure Black Typography** | Mọi subtitle, card header, caption mặc định dùng `#000000`. CẤM ép cứng màu xanh đen. |

### E. Document & Codebase Security Invariants (ERR_SEC_001–007)

| Mã Lỗi | Tên Bất Biến | Nguyên Tắc Bắt Buộc |
|---|---|---|
| `ERR_SEC_001` | **Macro & Malicious Code Execution** | Quét toàn diện `vbaProject.bin`, `vbaData.xml`, XLM macros và PDF `/Launch`, `/JavaScript`. Strip luồng macro nguy hiểm. |
| `ERR_SEC_002` | **Formula & DDE Injection (CSV/Excel)** | Ô bắt đầu bằng `=`, `+`, `-`, `@` gọi hàm nguy hại phải hóa giải bằng tiền tố nháy đơn `'`. |
| `ERR_SEC_003` | **Dangerous OLE Package Embeddings** | Quét `embeddings/oleObject*.bin`, phát hiện header PE (`MZ`) hoặc phần mở rộng `.exe`, `.bat`, `.ps1`. |
| `ERR_SEC_004` | **Hidden PII & Sensitive Metadata Leakage** | Xóa metadata tác giả (`dc:creator`, `cp:lastModifiedBy`), ghi chú nội bộ, và phục hồi sheet `veryHidden`. |
| `ERR_SEC_005` | **Safe XML & Zip Slip Defense** | Chống path traversal `../` trong tệp ZIP; vô hiệu hóa DTD external entities (`<!ENTITY`, `SYSTEM`) chống XXE. |
| `ERR_SEC_006` | **Deep CDR Promotion & Rels Clean** | Bóc gỡ `vbaProject.bin`/OLE, xóa `<Override>` và `<Relationship>` mồ côi, nâng cấp `macroEnabled` về OpenXML sạch. |
| `ERR_SEC_007` | **Vietnamese PII Redactor Financial Guard** | CCCD 12 số kiểm tra mã tỉnh `001..096`; SĐT kiểm tra đầu mạng VN; STK ngân hàng bắt buộc có Contextual Anchoring. |
