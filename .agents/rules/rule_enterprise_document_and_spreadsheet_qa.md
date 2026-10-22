# Rule: Enterprise Document & Spreadsheet QA / Diagnostics Standard

> **Scope**: Áp dụng bắt buộc cho toàn bộ các tác vụ đọc, sửa, tạo, chẩn đoán (QA/Diagnostic) và chuyển đổi tài liệu văn phòng
> (`.docx`, `.xlsx`, `.pdf`, `.md`) cùng hệ thống sơ đồ kỹ thuật / biểu đồ (Diagrams & Charts).

---

## 1. Triết lý Cốt lõi: Template-First & Non-Destructive Mutation

1. **Tuyệt đối không hardcode định dạng**: Không gán cứng mã màu hex, tên font hay kích thước ô nếu không đối chiếu từ template.
2. **Kế thừa dòng mẫu (Prototype Row Cloning)**: Luôn lấy một dòng dữ liệu đại diện có sẵn trong template làm khuôn mẫu để clone 100% style sang các dòng mới.
3. **An toàn khi sửa đổi (Zero-Destruction Invariant)**: Không bao giờ ghi đè trực tiếp lên file gốc nếu chưa tạo bản sao lưu dự phòng (`.bak`). Mặc định xuất bản sửa lỗi ra file `[name]_repaired.[ext]`.
4. **Phát hiện mốc ngữ nghĩa (Semantic Anchor Discovery)**: Dò tìm dòng Header, Data Start và Summary/Total bằng từ khóa và công thức, không dựa vào số dòng cố định.

---

## 2. Tiêu chuẩn Invariants 4 Domain Kỹ thuật

### A. Domain EXCEL (XLSX) — Cấu trúc & Định dạng Bảng tính
* **`ERR_XLSX_001` (Prototype Row Style Cloning)**: Khi thêm dữ liệu vào bảng, sao chép toàn bộ thuộc tính (`font`, `fill`, `border`, `alignment`, `number_format`) từ dòng dữ liệu đầu tiên ngay dưới Header.
* **`ERR_XLSX_002` (Dynamic Formula Range Expansion)**: Quét công thức gốc có sẵn trong template, dùng Regex bóc tách dải ô tham chiếu (ví dụ: `F10:F25`) và tịnh tiến động thành `F10:F{25 + \Delta K}` khi phình to dòng.
* **`ERR_XLSX_003` (Semantic Anchor Discovery)**: Dữ liệu mới CHỈ ĐƯỢC CHÈN vào giữa Data Start (dòng dưới Header) và Summary Sentinel (dòng xuất hiện `Total`, `Summary`, hoặc hàm `=SUM`).
* **`ERR_XLSX_004` (Safe Merged-Cell Introspection)**: Duy trì helper `get_effective_cell()`. Chỉ ghi giá trị vào ô Top-Left; đồng bộ style cho toàn bộ các ô trong dải merged để tránh rách viền ô.
* **`ERR_XLSX_005` (Adaptive 2D Freeze Panes & Dynamic Auto-scaling)**: Freeze Panes đặt tại giao điểm `(Header Row + 1, Cột dữ liệu đầu tiên)`. Chiều cao dòng tính tự động:
  $$\text{Row Height} = \max(\text{min\_h}, \text{total\_lines} \times \text{line\_h}) \quad \text{kết hợp} \quad \text{wrap\_text=True}$$
* **`ERR_XLSX_006` (DrawingML & Package Integrity)**: Luôn load trực tiếp template gốc để sửa đổi (`load_workbook(..., data_only=False)`). Tuyệt đối không tạo mới bằng `Workbook()` rỗng làm mất logo và shapes ở Cover sheet.
* **`ERR_XLSX_007` (Data Type Coercion & Leading-Zero Loss)**: Cột mang ngữ nghĩa định danh (ID, Code, Mã) bắt buộc gán `number_format = '@'` và truyền kiểu chuỗi tường minh.
* **`ERR_XLSX_008` (Print-Ready Page Setup)**: Thiết lập in ấn chuẩn A4 Landscape, `fitToPage = True`, `fitToWidth = 1`, `fitToHeight = 0` chống tràn lề trang in.
* **`ERR_XLSX_009` (15-Digit Precision & Accounting Negative Guard)**: CCCD (12 số), thẻ ngân hàng (16-19 số), số tài khoản ép kiểu `str` và gán `number_format = '@'`. Parse ngày tiếng Việt sang `datetime.date` + `'DD/MM/YYYY'`. Chuyển số âm `(100.000)` thành float âm `-100000`, gạch ngang `-` thành `0` với format kế toán. Chuẩn hóa Locale thập phân VN `20.000,50` $\rightarrow$ `20000.50`.
* **`ERR_XLSX_010` (XML Illegal Character Sanitization & 32k Limit)**: Khử sạch ký tự điều khiển ASCII `[\x00-\x08\x0b\x0c\x0e-\x1f]`. Cắt tỉa chuỗi tối đa 32.000 ký tự kèm hậu tố `... [Xem tiếp tài liệu gốc]`.
* **`ERR_XLSX_011` (_xlfn Modern Formula Compatibility & Reference Shifting)**: Tự động thêm tiền tố `_xlfn.` cho các hàm OpenXML mới (`XLOOKUP`, `CONCAT`, `TEXTJOIN`, `UNIQUE`, `FILTER`, `SORT`, `SORTBY`, v.v.) chống lỗi `#NAME?` trên Excel 2016/2019. Duy trì Regex AST Reference Shifting (`ERR_XLSX_002`).
* **`ERR_XLSX_012` (Gridlines, AutoFilter & Unlocked Row Height)**: Bật hiển thị đường lưới `ws.views.sheetView[0].showGridLines = True`. Tự động gắn vùng lọc `ws.auto_filter.ref`. Để `ws.row_dimensions[r].height = None` cho dòng dữ liệu cho phép Excel tự co giãn theo dòng `\n`.
* **`ERR_XLSX_013` (Selective Formula & Structure Protection)**: Mọi ô chứa công thức (`=...`) và Header mặc định khóa `locked=True`. Chỉ mở khóa `locked=False` cho vùng nhập liệu được chỉ định trước khi kích hoạt `ws.protection.enable()`.
* **`ERR_XLSX_014` (RAM-Constrained Streaming Threshold)**: Tự động chuyển đổi sang `read_only=True` khi đọc và `write_only=True` khi ghi nếu kích thước tệp $> 10\text{MB}$ hoặc số dòng $> 30.000$, đảm bảo trần RAM $\le 50\text{MB}$.

### B. Domain WORD (DOCX) — Tính Toàn vẹn OpenXML & Bảng Biểu
* **`ERR_DOCX_001` (The Last Paragraph Rule)**: Mọi cell trong bảng (`<w:tc>`) bắt buộc phải kết thúc bằng tối thiểu một đoạn văn (`<w:p>`). Không bao giờ để ô trống rỗng làm Word báo file hỏng.
* **`ERR_DOCX_002` (Cell Text Overwrite Stripping Formatting)**: Không gán `cell.text = "..."`. Phải thao tác qua `cell.paragraphs[0].runs` để giữ nguyên font chữ, cỡ chữ và màu sắc của template.
* **`ERR_DOCX_003` (Multi-Page Table Pagination Rupture)**: Bảng dài nhiều trang bắt buộc có `<w:cantSplit/>` trên từng dòng và `<w:tblHeader/>` trên hàng tiêu đề để tự động lặp lại ở mọi trang.
* **`ERR_DOCX_004` (Table Cell Shading & Border XML Desync)**: Sử dụng cú pháp có namespace đầy đủ (`parse_xml(f'<w:shd {nsdecls("w")} .../>')`) khi can thiệp XML bảng.
* **`ERR_DOCX_005` (Inline Shapes Aspect Ratio & Margin Overflow)**: Khóa tỉ lệ ảnh và giới hạn chiều rộng hình ảnh không vượt quá vùng in khả dụng:
  $$\text{Max Image Width} = \text{page\_width} - \text{left\_margin} - \text{right\_margin}$$
* **`ERR_DOCX_006` (Template Style Inheritance)**: Ưu tiên kế thừa Table Style có sẵn trong `styles.xml` của template (`table.style = '...'`) thay vì inject XML thô tràn lan.

### C. Domain DIAGRAMS & CHARTS — Chế độ Kép (Template vs Zero-Template)
* **`ERR_DIAG_001` (Template-Driven Token Drift)**: Khi có template, trích xuất 100% token kiểu dáng (kích thước, màu sắc, font chữ, viền) làm căn cứ duy nhất.
* **`ERR_DIAG_002` (Zero-Template Layout Collision)**: Khi vẽ từ đầu không có template:
  - Áp dụng nguyên tắc phối màu **60-30-10** (60% nền trung tính, 30% cấu trúc thẻ trắng/slate, 10% màu nhấn ngữ nghĩa).
  - Bố trí node trên Canvas Grid có khoảng cách tối thiểu 40px giữa các box và bo tròn góc 8px.
* **`ERR_DIAG_003` (Chart Anchor Overlap)**: Khi mở rộng dòng trong bảng dữ liệu, tự động relink chuỗi formula của biểu đồ (`val.numRef.f`) và dời tọa độ `chart.anchor._from.row` đặt ngay dưới bảng tổng hợp.

### D. Domain CONVERSION PIPELINE — Chuyển đổi Bền vững
* **`ERR_CONV_001` (Headless Conversion Layout Drift)**: Ưu tiên bộ lọc `writer_pdf_import` của LibreOffice/Stirling-PDF cho tiến trình PDF $\rightarrow$ DOCX.
* **`ERR_CONV_002` (Font Embedding & Missing Glyphs)**: Chỉ định chuỗi font fallback đầy đủ: `'Segoe UI, -apple-system, BlinkMacSystemFont, Roboto, Arial, sans-serif'`.
* **`ERR_CONV_003` (Markdown-Office Roundtrip Semantic Loss)**: Sử dụng tầng trung gian JSON AST để bảo toàn cấu trúc bảng phức tạp.
* **`ERR_CONV_004` (Zombie Lock Files Cleanup)**: Context Manager tự động thu gom sạch các file rác khóa hệ thống (`.~lock.*` và `~$*`) trước và sau khi thực thi.

---

## 3. Tiêu chí Nghiệm thu Định lượng (Quantitative Benchmark)

Một tài liệu chỉ được coi là đạt chuẩn xuất xưởng khi vượt qua bài kiểm tra chẩn đoán với:
1. **Excel**:
   - Số lượng DrawingML shapes: Giữ nguyên 100% (không bị mất ảnh/logo ở Cover).
   - Zero MergedCell write violations (0 exception).
   - 100% công thức live (không có `#REF!`, `#VALUE!`, `#DIV/0!`).
2. **Word**:
   - Zero Corrupt XML (100% cell bảng có `<w:p>`).
   - 100% bảng nhiều trang có `<w:tblHeader/>` và `<w:cantSplit/>`.
   - Chiều rộng bảng và ảnh không tràn lề trang giấy.
3. **CI/CD Exit Codes**:
   - `0`: Sạch lỗi hoàn toàn (Clean / Ready for Delivery).
   - `1`: Cảnh báo thẩm mỹ nhẹ (Warning: padding hẹp, wrap_text).
   - `2`: Lỗi cấu trúc nghiêm trọng (Critical: gãy công thức, hỏng XML, mất ảnh).

---

## 4. Quy trình Thực thi Lệnh Chuẩn (CLI Tool)

```bash
# Chẩn đoán một file Excel hoặc Word
python tool/pdf_to_docx_converter/tools/unified_qa_diagnostic.py --xlsx <path_to_file.xlsx>
python tool/pdf_to_docx_converter/tools/unified_qa_diagnostic.py --docx <path_to_file.docx>

# Chẩn đoán toàn bộ thư mục và xuất JSON log cho CI/CD
python tool/pdf_to_docx_converter/tools/unified_qa_diagnostic.py --all <dir_path> --format json --output report.json

# Tự động sửa lỗi cấu trúc an toàn (tạo file _repaired)
python tool/pdf_to_docx_converter/tools/unified_qa_diagnostic.py --xlsx <file.xlsx> --mode audit-fix

# Sửa lỗi in-place có sao lưu .bak
python tool/pdf_to_docx_converter/tools/unified_qa_diagnostic.py --docx <file.docx> --mode audit-fix --in-place
```
