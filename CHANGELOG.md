# CHANGELOG — Lịch Sử Thay Đổi & Tiến Hóa Dự Án (v1.0 → v3.4)

Tài liệu này tổng hợp toàn bộ lộ trình nâng cấp, các phân hệ kiến trúc mới, và các cải tiến công nghệ từ phiên bản cơ sở (v1.0 Baseline) cho đến phiên bản hoàn thiện hiện tại (**v3.4 Production Grade**) của bộ công cụ **`antigravity-doc-handler`**.

---

## 1. Bảng So Sánh Toàn Diện (Feature Comparison Matrix)

| Phân hệ / Tiêu chí kỹ thuật | Phiên bản cũ (v1.0 Baseline) | Phiên bản hiện tại (v3.4 Production Grade) | Lợi ích & Độ tin cậy thực tế |
|---|---|---|---|
| **Danh Sách & Đánh Số Word (Lists & Numbering)** | Dùng style `List Number` tĩnh của Word khiến số bị đếm tích lũy liên tục không reset; bullet không phân cấp; không hỗ trợ checkbox | **Native OpenXML Numbering Engine (`docx_lists.py`)**: Cấp phát `numId` động tự động reset về 1 (`startOverride=1`); Smart Continuation duy trì chuỗi số cha khi có sub-items xen kẽ; phân cấp 3 tầng (1->a->i, •->○->■); hỗ trợ checkbox (`☐`, `☑`); tương thích cả trong Table cells | Khi mở tài liệu trên Microsoft Word, nhấn Enter tự động sinh số tiếp theo; danh sách mới tự động bắt đầu từ 1; tuyệt đối không sinh số bừa bãi cho text thường |
| **Lưu Trữ Lịch Sử Tệp (History & Storage)** | Không có; chuyển đổi xong phải lưu thủ công | **Zero-Config SQLite Storage (`db_manager.py`)**: Tự động lưu ngầm lịch sử chuyển đổi vào `.web_outputs/studio_history.db`, tích hợp Drawer trượt trên Web Studio | 1-click tải lại tệp cũ hoặc xóa lịch sử; zero phụ thuộc bên ngoài; bảo vệ hoàn toàn file cá nhân trên Git |
| **Soạn Thảo Toán Học Trực Quan (Math Studio)** | Không có giao diện; chỉ chạy script CLI | **Interactive Math Studio & Equation Lab**: Palette chèn ký hiệu nhanh, mẫu công thức chuẩn kỹ thuật, Live Preview thời gian thực, 1-click xuất tệp Word DOCX | Soạn thảo công thức LaTeX trực quan trên trình duyệt và xuất sang Word chuẩn OMML với định dạng chuyên nghiệp |
| **Xử lý Hình ảnh DOCX** | Chỉ ghi text giữ chỗ dạng `[IMAGE: path]` | Nhúng ảnh nhị phân thực tế (Local path, Asset directory, Base64), tự động co ảnh bảo vệ lề in (`ERR_DOCX_005`), có thẻ Fallback trực quan khi thiếu ảnh | Tài liệu sinh ra có ảnh sắc nét Retina 300+ DPI, 100% không vỡ bố cục khi in ấn A4 hoặc xuất PDF |
| **Bảng biểu DOCX (Tables)** | Chỉ hỗ trợ ma trận ô phẳng $1 \times 1$; không hỗ trợ gộp ô; bảng dài tràn trang bị vỡ nét viền | Hỗ trợ gộp ô đa chiều (`colspan` ngang, `rowspan` dọc); tự động chèn `<w:cantSplit/>`, `<w:tblHeader/>`, `<w:vAlign w:val="center"/>`; tuân thủ tuyệt đối *The Last Paragraph Rule* (`ERR_DOCX_001`) | Bảng biểu hiển thị chuẩn OpenXML ISO/IEC 29500; tiêu đề tự lặp lại ở mọi trang; không bao giờ bị cắt đôi hàng chữ ngang mép trang |
| **Công thức Toán Word (DOCX Math)** | Không hỗ trợ; text công thức bị trơ hoặc lỗi phông | **Native Math Expression Engine (`docx_math.py`)**: Kiến trúc Dual-Mode kết hợp OpenXML Multi-Run (`<w:vertAlign>`, `<w:i/>`, Greek lexicon) cho inline `$ ... $` và OMML (`<m:oMath>`, `<m:f>`) cho phân số/block equation | Hiển thị công thức toán học chuyên nghiệp, ký tự Hy Lạp ($\alpha, \beta, \omega, \Delta$) và ký số dưới ($v_{ref}, P_{max}$) tương thích 100% mọi trình xem (Word, LibreOffice, WPS, Google Docs, Mobile) |
| **Bảo tồn Mẫu Word (Template Patching)** | Không có; chỉ ghi đè tạo file mới từ đầu làm mất Trang bìa và Header/Footer doanh nghiệp | **Chế độ In-Place Template Patching (`patch_docx_template`)**: Quét và thay thế placeholder `{{KEY}}` (hợp nhất run chống vỡ token) và `{{BLOCK:id}}` (chèn bảng/ảnh động) | Giữ nguyên 100% Trang bìa (Cover Page), DrawingML, Logo doanh nghiệp, Header, Footer và mục lục tự động |
| **Trình diễn PowerPoint (PPTX Engine)** | Không hỗ trợ (chỉ có chuyển đổi PDF/Word cơ bản) | **Bộ sinh Slide 16:9 Tự động (`pptx_writer.py`, `pptx_engine.py`)**: Tỷ lệ chuẩn Widescreen 16:9 (`PPTX_INV_01`), Typography thuần đen `#000000` (`PPTX_INV_10`), Run-level formatting supremacy (`PPTX_INV_09`) | Tạo bài thuyết trình kỹ thuật chuẩn 1080p từ Markdown hoặc JSON AST; Ribbon và Color Picker của PowerPoint hoàn toàn tự do, không bị khóa cứng định dạng |
| **Toán học PowerPoint (PPTX Math)** | Không hỗ trợ | **Native DrawingML Math (`pptx_math.py`)**: Subscript (`baseline="-25000"`), Superscript (`baseline="30000"`), biến in nghiêng, đơn vị đứng thẳng | Công thức toán học và ký hiệu vật lý ($v_{ref} = 0\text{ m/s}$) hiển thị chuẩn DrawingML gốc, 0 lỗi namespace |
| **Xử lý Bảng tính Excel (Universal XLSX)** | Không hỗ trợ xử lý nâng cao, dễ mất hình khối DrawingML | **Universal Excel Engine (`xlsx_writer.py`)**: Chế độ Native openpyxl bảo toàn hình vẽ/ảnh, dịch chuyển công thức động qua Regex (`ERR_XLSX_002`), clone kiểu dáng từ prototype row (`ERR_XLSX_001`) | Cho phép điền dữ liệu động vào mẫu báo cáo tài chính/kỹ thuật phức tạp mà không làm hỏng logo, biểu đồ hoặc công thức KPI |
| **Sơ đồ Kỹ thuật (Diagram Suite)** | Chỉ có sơ đồ cơ bản qua Chromium headless | **Hệ sinh thái Đa Engine**: Hỗ trợ 4 loại công nghệ: Draw.io Native mxGraph (`mxgraph_engine.py`), PlantUML cục bộ (`plantuml.jar` với chuẩn C4 nội bộ), Mermaid pastel phẳng, và Canvas Screen Flow tương tác | Xuất vector SVG và PNG siêu nét 300+ DPI, hỗ trợ kéo thả chỉnh sửa 2 chiều qua `diagram_editor.py` |
| **Quản lý Định dạng (Decoupled Styling)** | Code và Style bị trộn lẫn, khó thay đổi giao diện | Tách rời hoàn toàn Nội dung (`.md`) và Kiểu dáng (`.style.yaml`); tự động trích xuất metadata phong cách từ DOCX mẫu | Thay đổi toàn bộ phông chữ, màu sắc, lề trang chỉ bằng việc chỉnh sửa file YAML mà không cần sửa code |
| **Giao diện & CLI Dispatcher** | Các script chạy phân mảnh | `ai_tools_cli.py` hợp nhất toàn bộ lệnh: `docx-tools`, `pptx-build`, `pptx-preview`, `diagram-render`, `convert`, `inspect-doc` + Desktop GUI Tkinter (`main.py`) | Thân thiện cho cả người dùng cuối (chạy app desktop) và AI Coding Agent (gọi CLI dạng declarative JSON) |
| **Kiểm định Chất Lượng (Quality Assurance)** | 16 bài test sơ khai (bị lỗi hardcode đường dẫn trên máy khác) | **105 Automated Test Suites (100% Pass Rate)**: Bộ test toàn diện bao quát OpenXML Lists, Invariants, DrawingML, SQLite CRUD, Web API, và Microsoft Word COM Automation | Bảo đảm tính ổn định tuyệt đối trong môi trường CI/CD và máy tính cá nhân |

---

## 2. Chi Tiết Các Phiên Bản Phát Triển

### Phiên Bản 3.4.0 (Phiên Bản Hiện Tại — OpenXML List & Numbering Engine + 105 Test Suites)
- **Động cơ Quản lý Danh sách Độc lập (`docx_lists.py`)**:
  - Đăng ký cấu trúc OpenXML `abstractNum` chuẩn ECMA-376 (ID 2001 cho số thứ tự 3 cấp `1.` -> `a.` -> `i.`; ID 2002 cho bullet 3 cấp `•` -> `○` -> `■`).
  - Cấp phát `numId` động độc lập kèm `<w:startOverride w:val="1"/>`, tự động reset danh sách về 1 khi bắt đầu mục mới.
  - **Cơ chế Smart Continuation**: Duy trì liên tục số thứ tự cấp cha ngay cả khi có ghi chú hoặc sub-bullet xen kẽ.
  - **Dynamic Native Word Numbering**: Khi mở tài liệu trong Word, nhấn Enter sẽ tự động nhảy số tiếp theo thay vì văn bản tĩnh.
  - **Task Lists Checkbox**: Chuyển đổi mượt mà `- [ ]` sang `☐` (ballot box) và `- [x]` / `* [X]` sang `☑` (checked box) in đậm.
  - **Hanging Indentation Invariants**: Áp dụng chuẩn thụt lề dxa `<w:ind w:left="..." w:hanging="360"/>` (cấp 0: 720, cấp 1: 1440, cấp 2: 2160).
  - **Strict Guard**: Tuyệt đối không tự ý sinh số thứ tự hoặc bullet cho các đoạn văn thông thường.
- **Tích hợp sâu vào `markdown_converter.py` & `docx_writer.py`**:
  - `markdown_converter.py`: Tích hợp `ListContextTracker` nhận diện thụt đầu dòng; tự động reset danh sách khi gặp tiêu đề `#`, code blocks, ảnh, bảng biểu, ngắt trang, hoặc văn bản thường.
  - **Hỗ trợ Danh sách trong Bảng (Table Cell Lists)**: Phân giải các dòng ngắt `<br>` hoặc `\n` trong ô bảng biểu thành các đoạn danh sách chuẩn OpenXML, giữ vững quy tắc *The Last Paragraph Rule* (`ERR_DOCX_001`).
  - `docx_writer.py`: Hỗ trợ trường `list_type`, `list_level`, `list_restart`, `is_checked` trong paragraph AST và cell paragraphs.
- **Mở Rộng Test Suites Lên 105 Bài Kiểm Thử (100% Pass Rate)**:
  - Tạo mới `tests/test_docx_lists.py` (22 bài test chuyên sâu cho OpenXML list engine).
  - Tạo mới `tests/test_docx_extended.py` (25 bài test mở rộng kiểm tra các bất biến bố cục, màu sắc, bảng biểu, lề trang, Math, và template patching).
  - Tổng số bài test tăng từ 58 lên **105 bài kiểm thử tự động**, đạt tỷ lệ đạt 100/100 (100% Pass Rate).

### Phiên Bản 3.3.0 (End-User Suite, Storage & Diagram Precision)
- **Tích hợp SQLite Storage Layer (`db_manager.py`)**:
  - Tự động ghi nhận lịch sử các tác vụ chuyển đổi tệp: `convert`, `docx_to_pptx`, `pdf_to_pptx`, `diagram_render`, `math_export`.
  - Database được đặt tại `.web_outputs/studio_history.db` (được `.gitignore` bảo vệ, zero nguy cơ rò rỉ dữ liệu lên Git).
  - Cung cấp trọn bộ CRUD: `record_history`, `get_recent_history`, `delete_history_item`, `clear_history`.
- **Slide-out History Drawer trên Giao Diện Web**:
  - Bổ sung nút "Lịch Sử" kèm badge số lượng tệp xử lý trên Header.
  - Ngăn kéo trượt (Slide-out Drawer) với hiệu ứng kính mờ (glassmorphism), hỗ trợ 1-click tải lại file hoặc xóa từng mục/xóa sạch lịch sử.
- **Xây dựng Math Studio & Equation Lab**:
  - Giao diện trực quan cho phép soạn thảo công thức toán học LaTeX với palette ký hiệu nhanh ($\alpha, \beta, \omega, \Delta, \approx, \le, \ge, \pm, \frac{a}{b}, \sqrt{x}, x^2, v_{ref}$).
  - Cung cấp sẵn các bộ mẫu công thức kinh điển (Động học Robot, Vật lý/Điện áp, Giải tích giới hạn, Hình học).
  - Trình xem trước thời gian thực (Live Preview) phân giải phân số, căn thức, chỉ số trên/dưới.
  - 1-click xuất tệp Word DOCX chuẩn OMML và OpenXML qua endpoint `/api/math/export`.
- **Tối ưu hóa & Khắc phục Triệt để Engine Sơ Đồ Mermaid.js**:
  - **Triệt tiêu lỗi dẹt khung hình (Anti-Pancake Aspect Ratio)**: Gỡ bỏ triệt để cấu hình hardcode `"width": 2048` tại cấp root trong `clean_modern.json` và `enterprise.json`, giúp Mermaid CLI tự động tính toán bounding box chính xác theo chiều cao thực của lifelines thay vì nén dẹt sơ đồ xuống dải mỏng 90px (tỷ lệ dị dạng 17:1).
  - **Chuẩn hóa Typography & Inter Font**: Chuyển đổi toàn bộ phông chữ sơ đồ sang font sans-serif hiện đại `Inter, system-ui, -apple-system`, thay thế font monospace `JetBrains Mono` thô rộng, giúp các hộp actor thu gọn thanh thoát và sắc nét hơn.
  - **Ngắt dòng nhãn hành động (Multi-line Action Wrapping)**: Áp dụng chuẩn [Rule 8 & 9](file:///d:/New%20folder%20%284%29/antigravity-doc-handler-main/.agents/rules/rule_technical_diagram_standards.md) chèn thẻ `<br/>` cho tất cả nhãn dài (>20 ký tự), phục hồi tỷ lệ khung hình chuẩn vàng **1.57 : 1** (`1888 × 1204 px`) hoàn hảo cho in ấn A4 và slide 16:9.
  - **Cải tiến Preview UX & Vite Watcher**: Đổi mặc định sau render sang chế độ `Khớp màn hình` (`mode-fit`) để người dùng có cái nhìn bao quát ngay lập tức; đồng thời cấu hình `server.watch.ignored` trong `vite.config.js` loại trừ các thư mục `.web_outputs/`, `diagram_assets/` nhằm loại bỏ triệt để hiện tượng trình duyệt tự reload toàn trang khi render ảnh.
- **Bổ sung kiểm thử `tests/test_history_and_db.py`**:
  - Kiểm tra toàn diện hoạt động của SQLite CRUD và các HTTP endpoints.
  - Nâng tổng số test case toàn dự án lên **58 tests (100% Passed)**.

### Phiên Bản 3.2.0
- **Bổ sung Native Math Expression Engine cho Word DOCX (`docx_math.py`)**:
  - Xây dựng bộ Lexer & Tokenizer phân tích biểu thức LaTeX toán học.
  - Chuyển đổi linh hoạt giữa OpenXML Multi-Run (cho `$ ... $`) và OMML Office Math (cho `$$ ... $$` và phân số `\frac{a}{b}`).
  - Biến số tự động in nghiêng (`<w:i/>`), text và đơn vị đo lường đứng thẳng (`<w:i w:val="0"/>`), chỉ số dưới/trên dùng thẻ OpenXML `<w:vertAlign/>`.
- **Tích hợp sâu vào Core DOCX & Markdown**:
  - `docx_writer.py`: Tự động phát hiện và render công thức toán học trong đoạn văn, run-level data, và toàn bộ cell của bảng biểu.
  - `markdown_converter.py`: Trình biên dịch Markdown sang DOCX tự động nhận diện cú pháp toán inline và căn giữa các khối phương trình toán học block.
- **Bổ sung kiểm thử `tests/test_docx_math.py`**:
  - Đạt 10/10 bài kiểm thử đơn vị và tích hợp.
  - Xác thực qua Microsoft Word COM Automation (mở tài liệu ở chế độ headless không sinh bất kỳ hộp thoại cảnh báo hay lỗi cấu trúc).
- **Nâng tổng số test case toàn dự án lên 56 tests (100% Passed)**.

### Phiên Bản 3.1.0
- **Bổ sung Native DrawingML Math Engine cho PowerPoint (`pptx_math.py`)**:
  - Xây dựng bộ token hóa biểu thức toán cho slide thuyết trình.
  - Ứng dụng native DrawingML Subscript/Superscript thông qua thuộc tính baseline XML.
- **Hoàn thiện bộ bất biến PowerPoint Presentation (PPTX_INV_01 → PPTX_INV_13)**:
  - Khóa tỷ lệ khung hình 16:9 Widescreen (`PPTX_INV_01`).
  - Thiết lập Typography thuần đen `#000000` mặc định (`PPTX_INV_10`).
  - Đảm bảo quyền kiểm soát màu sắc và định dạng run-level (`PPTX_INV_09`).
  - Dọn dẹp triệt để AlternateContent ghost shapes (`PPTX_INV_13`).

### Phiên Bản 3.0.0
- **Phát triển Universal Excel Engine (`xlsx_writer.py`)**:
  - Cơ chế Native openpyxl giữ nguyên các đối tượng đồ họa DrawingML trên Sheet bìa (`ERR_XLSX_006`).
  - Thuật toán tịnh tiến công thức động qua Regular Expression (`ERR_XLSX_002`).
  - Sao chép toàn bộ định dạng từ dòng dữ liệu mẫu (Prototype Row Style Cloning `ERR_XLSX_001`).
  - Gán định dạng chuỗi `@` cho các cột mã định danh (`ERR_XLSX_007`).
- **Phát triển Native Draw.io XML Engine (`mxgraph_engine.py`)**:
  - Tuân thủ 20 nguyên tắc bất biến mxGraph (`MX_INV_01` → `MX_INV_20`).
  - Đi dây trực giao vuông góc (Orthogonal Perimeter Routing), cổng kết nối chu vi tỷ lệ vi phân, nhãn dây có mask trắng ôm khít chống đè chữ.

### Phiên Bản 2.0.0
- **Decoupled Architecture**: Tách rời tệp nội dung Markdown và stylesheet YAML.
- **OpenXML Table Repair Pipeline**: Tự động sửa lỗi tràn bảng Word bằng `<w:cantSplit/>` và `<w:tblHeader/>`.
- **Diagram Studio**: Tích hợp PlantUML cục bộ (với thư viện C4 nội bộ không phụ thuộc internet) và Mermaid pastel.

### Phiên Bản 1.0.0 (Bản Cơ Sở)
- Chuyển đổi cơ bản PDF sang DOCX và ngược lại qua `pdf2docx` và `pymupdf`.
- Chưa có nhúng ảnh thực tế, bảng biểu sơ khai, không hỗ trợ công thức toán học hoặc trình diễn slide.

---

## 3. Danh Mục Các Tiêu Chuẩn Kỹ Thuật (Engineering Invariants)

Dự án áp dụng nghiêm ngặt các bộ quy chuẩn bất biến sau trong toàn bộ mã nguồn:

### Bộ Tiêu Chuẩn Word OpenXML (DOCX)
- `ERR_DOCX_001 (The Last Paragraph Rule)`: Mọi cell bảng (`<w:tc>`) bắt buộc phải kết thúc bằng tối thiểu một thẻ `<w:p>`.
- `ERR_DOCX_002 (Run Text Overwrite)`: Thao tác nội dung thông qua danh sách runs, không gán đè trực tiếp text làm mất định dạng.
- `ERR_DOCX_003 (Multi-Page Table Flags)`: Bảng nhiều trang bắt buộc có `<w:cantSplit/>` và `<w:tblHeader/>`.
- `ERR_DOCX_005 (Printable Margin Overflow)`: Chiều rộng hình ảnh luôn được khóa tỉ lệ và $\le$ chiều rộng vùng in khả dụng của trang A4.

### Bộ Tiêu Chuẩn PowerPoint DrawingML (PPTX)
- `PPTX_INV_01`: Slide master mặc định xuất tỷ lệ 16:9 (`1280 x 720 pt`).
- `PPTX_INV_09`: Toàn bộ thuộc tính phông chữ, in đậm và màu sắc phải đặt ở `<a:rPr>`, nghiêm cấm khóa chết ở `<a:defRPr>`.
- `PPTX_INV_10`: Mọi văn bản con, card header và số trang mặc định dùng đen thuần `#000000`, trao toàn quyền đổi màu trên Ribbon.
- `PPTX_INV_12`: Ký hiệu toán học/vật lý dùng Native DrawingML Subscript (`baseline="-25000"`), biến in nghiêng, đơn vị đứng thẳng.
- `PPTX_INV_13`: Duyệt cây hình học dọn sạch các khối `<mc:AlternateContent>` cũ nằm ngầm để tránh ghost data.

### Bộ Tiêu Chuẩn Bảng Tính Excel (XLSX)
- `ERR_XLSX_001`: Luôn clone 100% style từ dòng dữ liệu đại diện trong template sang dòng mới chèn vào.
- `ERR_XLSX_002`: Bóc tách dải ô tham chiếu công thức bằng Regex và tịnh tiến động khi mở rộng dòng.
- `ERR_XLSX_006`: Bảo toàn 100% DrawingML (ảnh, logo, shape) bằng cách load template gốc trực tiếp qua openpyxl.

---

## 4. Báo Cáo Kiểm Định Chất Lượng (Quality Verification)

- **Tổng số ca kiểm thử**: 56 bài test tự động.
- **Trạng thái thực thi**: 56/56 bài test **PASSED 100%**.
- **Môi trường xác nhận**:
  - Python 3.10, 3.11, 3.12, 3.14 (Windows 64-bit).
  - Microsoft Word COM Automation (Kiểm tra mở tài liệu không lỗi, 0 repair warnings).
  - Microsoft PowerPoint Automation (Kiểm tra kết xuất slide, 0 warning dialogs).

---

## 5. Hướng Dẫn Sử Dụng Nhanh Qua CLI

```bash
# 1. Chuyển đổi Markdown sang Word DOCX (Hỗ trợ công thức toán $...$ và $$...$$)
python ai_tools_cli.py convert document.md -t docx --style doc.style.yaml

# 2. Tạo bài thuyết trình PowerPoint 16:9 từ tệp JSON Spec
python ai_tools_cli.py pptx-build slide_spec.json -o presentation.pptx

# 3. Xuất ảnh preview các slide PowerPoint để kiểm định
python ai_tools_cli.py pptx-preview presentation.pptx -o ./previews/ --slides 1,2,3

# 4. Render sơ đồ kỹ thuật đa Engine (Draw.io / PlantUML / Mermaid / Canvas)
python ai_tools_cli.py diagram-render spec.json -o diagram.png

# 5. Khởi chạy giao diện Desktop Studio (Tkinter)
python main.py
```
