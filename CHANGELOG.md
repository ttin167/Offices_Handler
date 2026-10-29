# CHANGELOG — Lịch Sử Thay Đổi & Tiến Hóa Dự Án (v1.0 → v4.6.0)

Tài liệu này tổng hợp toàn bộ lộ trình nâng cấp, các phân hệ kiến trúc mới, và các cải tiến công nghệ từ phiên bản cơ sở (v1.0 Baseline) cho đến phiên bản hoàn thiện hiện tại (**v4.6.0 Production Grade — Enterprise Omni-Pipeline, Cryptographic PAdES/PDF-A, Two-Way Excel-to-Brief, Deep CDR & AST Code-to-Diagram Suite**) của bộ công cụ **`antigravity-doc-handler`**.

---

## 1. Bảng So Sánh Toàn Diện (Feature Comparison Matrix)

| Phân hệ / Tiêu chí kỹ thuật | Phiên bản cũ (v1.0 Baseline) | Phiên bản hiện tại (v4.6.0 Production Grade) | Lợi ích & Độ tin cậy thực tế |
|---|---|---|---|
| **Đường Ống Đa Bước & 6 Phân Hệ v4.6 (Omni-Pipeline, Deep CDR, PAdES & Code-to-Diagram)** | Chạy rời rạc từng công cụ đơn lẻ; không hỗ trợ ký số mật mã PKCS#7, PDF/A-2b, Excel-to-Brief hay AST Code-to-Diagram | **6 Phân Hệ Chiến Lược v4.6.0 & 5 Chốt Chặn Edge-Case**: (1) Ký số mật mã PKCS#7/PAdES (`/Type /Sig`, SHA-256 digest, `.pfx`/`.p12`) & Lưu trữ dài hạn ISO 19005-2 **PDF/A-2b** (`pdfaid:part=2`, sRGB `OutputIntents`); (2) **Excel-to-Brief** (`XLSX -> DOCX/MD/PDF` kèm `ERR_XLSX_015` Dual-Workbook tự tính `=SUM/AVERAGE/MIN/MAX` khi `data_only=True` trả về `None`) & **Cross-Workbook Consolidation**; (3) **Deep CDR** (`ERR_SEC_006`: bóc gỡ `vbaProject.bin`/OLE `MZ`, tự động nâng cấp `macroEnabled.main+xml` về chuẩn OpenXML và dọn sạch `.rels`) & **Smart Vietnamese PII Redactor** (`ERR_SEC_007`: kiểm tra mã tỉnh CCCD `001..096`, đầu số nhà mạng VN không kèm `VND/USD/đ/%`, Contextual Anchoring `STK/Số tài khoản` chống nhầm số tiền BCTC); (4) **Declarative Workflow Pipeline** (`core/pipeline/`, `ERR_PIPE_001`: trần `MAX_JOB_WORKERS=2` cho Render 512MB RAM, dọn sạch 100% file tạm `step_*_temp.*` trong `finally`); (5) **AST Code-to-Diagram** (quét Python AST & SQL/Prisma DDL ra PlantUML/Mermaid/Draw.io + **Pure Vector SVG Exporter** + tự động chèn vào `.docx`/`.pptx`); (6) **Unified Desktop App Mode** (`ui/webview_launcher.py`, `ERR_UI_001`: cô lập `--user-data-dir` giữ sống server đúng vòng đời cửa sổ App) | Tự động hóa khép kín từ mã nguồn, bảng tính, hợp đồng đến ký số mật mã và lưu trữ vĩnh viễn; an toàn tuyệt đối trên cả máy trạm Windows lẫn container đám mây 512MB RAM |
| **Bảo Mật Tài Liệu & Cloudflare Audit (Document Security Suite)** | Không hỗ trợ kiểm tra tĩnh mã độc; file xuất có thể chứa macro độc hại hoặc DDE/Formula Injection | **Cloudflare Security Audit Harness & Document Sanitizer (`core/security/`)**: Tích hợp kép 2 cấp độ (Agentic Skill 6-phase + Core App Engine Gate 13); Mô hình Phản Biện Đối Nghịch (Adversarial Verifier) chia 3 trạng thái (`confirmed`, `needs_validation`, `rejected`); Khử 7 bất biến an ninh tài liệu: `ERR_SEC_001`..`007`; CLI `doc-security`, MCP tool `doc_security_audit`, REST API `/api/security/*`, và công tắc Gate 13 + Deep CDR + Smart PII trên Web Studio | Bảo vệ 100% doanh nghiệp và người dùng trước các cuộc tấn công khai thác qua tệp tin Office/PDF/CSV; khử hoàn toàn nguy cơ RCE và rò rỉ dữ liệu nhạy cảm mà không làm mất dữ liệu người dùng |
| **Kỷ Yếu Kỹ Thuật Số & Nhà In (Yearbook & Photobook Studio)** | Không hỗ trợ; phải dùng các phần mềm thiết kế đồ họa đóng gói thủ công | **Digital Yearbook & Pre-Press Photobook Studio (`core/yearbook/`)**: Trình xem 3D Flipbook 2 trang đôi mở phẳng (Layflat 180°); Rãnh gáy sách mô phỏng độ sâu giấy thật; Kéo thả hoán đổi ảnh (Slot Swapping); Công tắc Vạch Xén In Ấn (Bleed 3mm Đỏ, Trim Xanh dương, Safe Zone Xanh lá); Động cơ in công nghiệp 300 DPI kèm Crop Marks và nhúng font Unicode tiếng Việt (`ERR_PDF_FONT_001`); Chuẩn hóa xoay EXIF (`ERR_EXIF_001`); Cắt ảnh thông minh top-bias 35% (`ERR_IMG_CROP_001`); Gói bàn giao thợ thiết kế 5-trong-1 (.zip) kèm file Word Brief chuẩn OpenXML (`cantSplit`, `tblHeader`, `vAlign="center"`) | Giải quyết bài toán kỷ yếu cử nhân tốt nghiệp và album ảnh số lượng lớn (50–100+ ảnh) từ trải nghiệm xem lật trang sống động trên Web đến tệp xuất xưởng gửi thẳng nhà in và thợ dàn trang |
| **Bảo Mật & Xử Lý PDF (PDF Pro Studio)** | Chỉ có 5 công cụ cơ bản (merge, split, rotate, watermark, compress) | **Bộ Công Cụ PDF Pro Studio & Visual Organizer**: True Redaction xóa vĩnh viễn byte/pixel (`ERR_PDF_001`); Mã hóa AES-256 phân quyền in/copy/modify; Đóng dấu giáp lai đa trang $N$-lát cắt overlap 8% (`ERR_PDF_002`); Chuẩn hóa A4 vector centering & số Bates Unicode (`ERR_PDF_003`); Trích xuất bảng sang Excel giữ số 0 đầu (`ERR_XLSX_007`) & khổ in A4 ngang (`ERR_XLSX_008`); Trình biên tập trang trực quan (xoay 90°, xóa trang, thumbnail 120 DPI) | Biến phân hệ PDF thành vũ khí cấp Stirling-PDF / Adobe Acrobat Pro; giải quyết trọn vẹn yêu cầu bảo mật thông tin tài chính/hồ sơ thầu và chuẩn hóa hợp đồng doanh nghiệp Việt Nam |
| **Danh Sách & Đánh Số Word (Lists & Numbering)** | Dùng style `List Number` tĩnh của Word khiến số bị đếm tích lũy liên tục không reset; bullet không phân cấp; không hỗ trợ checkbox | **Native OpenXML Numbering Engine (`docx_lists.py`)**: Cấp phát `numId` động tự động reset về 1 (`startOverride=1`); Smart Continuation duy trì chuỗi số cha khi có sub-items xen kẽ; phân cấp 3 tầng (1->a->i, •->○->■); hỗ trợ checkbox (`☐`, `☑`); tương thích cả trong Table cells | Khi mở tài liệu trên Microsoft Word, nhấn Enter tự động sinh số tiếp theo; danh sách mới tự động bắt đầu từ 1; tuyệt đối không sinh số bừa bãi cho text thường |
| **Lưu Trữ Lịch Sử Tệp (History & Storage)** | Không có; chuyển đổi xong phải lưu thủ công | **Zero-Config SQLite Storage (`db_manager.py`)**: Tự động lưu ngầm lịch sử chuyển đổi vào `.web_outputs/studio_history.db`, tích hợp Drawer trượt trên Web Studio | 1-click tải lại tệp cũ hoặc xóa lịch sử; zero phụ thuộc bên ngoài; bảo vệ hoàn toàn file cá nhân trên Git |
| **Soạn Thảo Toán Học Trực Quan (Math Studio)** | Không có giao diện; chỉ chạy script CLI | **Interactive Math Studio & Equation Lab**: Palette chèn ký hiệu nhanh, mẫu công thức chuẩn kỹ thuật, Live Preview thời gian thực, 1-click xuất tệp Word DOCX | Soạn thảo công thức LaTeX trực quan trên trình duyệt và xuất sang Word chuẩn OMML với định dạng chuyên nghiệp |
| **Xử lý Hình ảnh DOCX** | Chỉ ghi text giữ chỗ dạng `[IMAGE: path]` | Nhúng ảnh nhị phân thực tế (Local path, Asset directory, Base64), tự động co ảnh bảo vệ lề in (`ERR_DOCX_005`), có thẻ Fallback trực quan khi thiếu ảnh | Tài liệu sinh ra có ảnh sắc nét Retina 300+ DPI, 100% không vỡ bố cục khi in ấn A4 hoặc xuất PDF |
| **Bảng biểu DOCX (Tables)** | Chỉ hỗ trợ ma trận ô phẳng $1 \times 1$; không hỗ trợ gộp ô; bảng dài tràn trang bị vỡ nét viền | Hỗ trợ gộp ô đa chiều (`colspan` ngang, `rowspan` dọc); tự động chèn `<w:cantSplit/>`, `<w:tblHeader/>`, `<w:vAlign w:val="center"/>`; tuân thủ tuyệt đối *The Last Paragraph Rule* (`ERR_DOCX_001`) | Bảng biểu hiển thị chuẩn OpenXML ISO/IEC 29500; tiêu đề tự lặp lại ở mọi trang; không bao giờ bị cắt đôi hàng chữ ngang mép trang |
| **Công thức Toán Word (DOCX Math)** | Không hỗ trợ; text công thức bị trơ hoặc lỗi phông | **Native Math Expression Engine (`docx_math.py`)**: Kiến trúc Dual-Mode kết hợp OpenXML Multi-Run (`<w:vertAlign>`, `<w:i/>`, Greek lexicon) cho inline `$ ... $` và OMML (`<m:oMath>`, `<m:f>`) cho phân số/block equation | Hiển thị công thức toán học chuyên nghiệp, ký tự Hy Lạp ($\alpha, \beta, \omega, \Delta$) và ký số dưới ($v_{ref}, P_{max}$) tương thích 100% mọi trình xem (Word, LibreOffice, WPS, Google Docs, Mobile) |
| **Bảo tồn Mẫu Word (Template Patching)** | Không có; chỉ ghi đè tạo file mới từ đầu làm mất Trang bìa và Header/Footer doanh nghiệp | **Chế độ In-Place Template Patching (`patch_docx_template`)**: Quét và thay thế placeholder `{{KEY}}` (hợp nhất run chống vỡ token) và `{{BLOCK:id}}` (chèn bảng/ảnh động) | Giữ nguyên 100% Trang bìa (Cover Page), DrawingML, Logo doanh nghiệp, Header, Footer và mục lục tự động |
| **Trình diễn PowerPoint (PPTX Engine)** | Không hỗ trợ (chỉ có chuyển đổi PDF/Word cơ bản) | **Bộ sinh Slide 16:9 Tự động (`pptx_writer.py`, `pptx_engine.py`)**: Tỷ lệ chuẩn Widescreen 16:9 (`PPTX_INV_01`), Typography thuần đen `#000000` (`PPTX_INV_10`), Run-level formatting supremacy (`PPTX_INV_09`) | Tạo bài thuyết trình kỹ thuật chuẩn 1080p từ Markdown hoặc JSON AST; Ribbon và Color Picker của PowerPoint hoàn toàn tự do, không bị khóa cứng định dạng |
| **Toán học PowerPoint (PPTX Math)** | Không hỗ trợ | **Native DrawingML Math (`pptx_math.py`)**: Subscript (`baseline="-25000"`), Superscript (`baseline="30000"`), biến in nghiêng, đơn vị đứng thẳng | Công thức toán học và ký hiệu vật lý ($v_{ref} = 0\text{ m/s}$) hiển thị chuẩn DrawingML gốc, 0 lỗi namespace |
| **Xử lý Bảng tính Excel (Universal XLSX)** | Không hỗ trợ xử lý nâng cao, dễ mất hình khối DrawingML | **Universal Excel Engine (`xlsx_writer.py`) & Formula Playbook (`formula_playbook.py`)**: Chế độ Native openpyxl bảo toàn hình vẽ/ảnh, dịch chuyển công thức động qua Regex (`ERR_XLSX_002`), clone kiểu dáng từ prototype row (`ERR_XLSX_001`), cẩm nang 32 hàm và sổ tính 2 sheets live calculations | Cho phép điền dữ liệu động vào mẫu báo cáo tài chính/kỹ thuật phức tạp mà không làm hỏng logo, biểu đồ hoặc công thức KPI; tra cứu và xuất cẩm nang công thức 4 chiều |
| **Sơ đồ Kỹ thuật (Diagram Suite)** | Chỉ có sơ đồ cơ bản qua Chromium headless | **Hệ sinh thái Đa Engine**: Hỗ trợ 4 loại công nghệ: Draw.io Native mxGraph (`mxgraph_engine.py`), PlantUML cục bộ (`plantuml.jar` với chuẩn C4 nội bộ), Mermaid pastel phẳng, và Canvas Screen Flow tương tác | Xuất vector SVG và PNG siêu nét 300+ DPI, hỗ trợ kéo thả chỉnh sửa 2 chiều qua `diagram_editor.py` |
| **Quản lý Định dạng (Decoupled Styling)** | Code và Style bị trộn lẫn, khó thay đổi giao diện | Tách rời hoàn toàn Nội dung (`.md`) và Kiểu dáng (`.style.yaml`); tự động trích xuất metadata phong cách từ DOCX mẫu | Thay đổi toàn bộ phông chữ, màu sắc, lề trang chỉ bằng việc chỉnh sửa file YAML mà không cần sửa code |
| **Giao diện & CLI Dispatcher** | Các script chạy phân mảnh | `ai_tools_cli.py` hợp nhất toàn bộ lệnh: `doc-security`, `yearbook-layout`, `yearbook-build`, `docx-tools`, `pptx-build`, `diagram-render`, `convert`, `inspect-doc` + Desktop GUI Tkinter (`main.py`) | Thân thiện cho cả người dùng cuối (chạy app desktop) và AI Coding Agent (gọi CLI dạng declarative JSON) |
| **Kiểm định Chất Lượng (Quality Assurance)** | 16 bài test sơ khai (bị lỗi hardcode đường dẫn trên máy khác) | **305 Automated Test Suites (100% Pass Rate)**: Toàn diện từ OpenXML, DrawingML, SQLite, Web REST/SSE, MCP Tools, PDF Pro Studio, Formula Playbook, Digital Yearbook Studio, đến Document Security Suite | Bảo đảm tính ổn định tuyệt đối trong môi trường CI/CD và máy tính cá nhân |

---

## 2. Chi Tiết Các Phiên Bản Phát Triển

### Phiên Bản 4.6.0 (Enterprise Omni-Pipeline, Cryptographic PAdES/PDF-A, Two-Way Excel-to-Brief, Deep CDR & AST Code-to-Diagram Suite)
- **Phân Hệ 1 — Chữ Ký Số Mật Mã PKCS#7 / PAdES & Lưu Trữ Dài Hạn PDF/A-2b (`core/converter/pdf_toolkit.py`)**:
  - `crypto_sign_pdf` & `verify_pdf_crypto_signature`: Nhúng từ điển chữ ký mật mã chuẩn ISO `/Type /Sig` (`/Filter /Adobe.PPKLite`, `/SubFilter /ETSI.CAdES.detached` hoặc `/adbe.pkcs7.detached`), băm toàn vẹn SHA-256 của tệp PDF, hỗ trợ chứng thư số `.pfx`/`.p12` qua `cryptography` hoặc con dấu xác thực nội bộ kèm khung chữ ký số trực quan trên trang đích.
  - `convert_to_pdfa`: Chuẩn hóa tài liệu PDF sang định dạng lưu trữ dài hạn quốc tế **ISO 19005-2 (PDF/A-2b / PDF/A-1b)**, tự động gỡ bỏ các thẻ động bị cấm (`/JavaScript`, `/JS`, `/Launch`, `/EmbeddedFiles`), nhúng gói siêu dữ liệu XMP RDF (`pdfaid:part=2`, `pdfaid:conformance=B`) và cấu hình không gian màu chuẩn sRGB IEC61966-2.1 vào `/OutputIntents`.
- **Phân Hệ 2 — Động Cơ Báo Cáo Bảng Tính 2 Chiều (`core/xlsx/reporting.py`) & Chốt Chặn `ERR_XLSX_015`**:
  - `excel_to_brief`: Tự động kết xuất bảng tính Excel (`.xlsx`) thành báo cáo tóm tắt điều hành sang **Word (`.docx`)**, **Markdown (`.md`)**, và **PDF (`.pdf`)** với bảng tổng hợp chỉ số KPI (`SUM`, `AVERAGE`, `MIN`, `MAX`, `COUNT`) và tuân thủ 100% bất biến bảng biểu OpenXML (`<w:cantSplit/>`, `<w:tblHeader/>`, `<w:vAlign w:val="center"/>`).
  - **Edge Case Guard #3 (`ERR_XLSX_015` — Un-cached Formula Fallback)**: Mở song song hai phiên bản Workbook (`data_only=True` và `data_only=False`); khi tệp Excel vừa được sinh tự động bằng `openpyxl` chưa có giá trị cache (`None`), bộ giải công thức nội bộ `_evaluate_basic_excel_formula` tự động tính toán chính xác `SUM`, `AVERAGE`, `COUNT`, `COUNTA`, `MIN`, `MAX` và biểu thức số học ô (`=B2-C2`, `=B2*C2`) để bảng báo cáo Word/PDF không bao giờ bị trống ô tổng cộng.
  - `consolidate_workbooks`: Hợp nhất hàng loạt tệp Excel chi nhánh theo 2 chế độ `append_rows` (gộp dòng kèm cột truy vết nguồn `Source_Workbook`, bảo vệ số 0 đầu `@` và dòng công thức tổng động) và `merge_sheets` (tổng hợp đa sheet kèm trang điều hướng `Index_Dashboard`).
- **Phân Hệ 3 — Deep CDR (Content Disarm & Reconstruction) & Smart Vietnamese PII Redactor (`core/security/sanitizer.py`)**:
  - **Edge Case Guard #1 (`ERR_SEC_006` — Macro-Enabled Content-Type Promotion & Rels Cleaning)**: Khi bóc gỡ `vbaProject.bin`, `vbaData.xml`, `/macrosheets/` hoặc gói nhị phân thực thi PE (`MZ`) `oleObject*.bin` khỏi tệp `.docm`/`.xlsm`/`.pptm`, `DocumentSanitizer` tự động xóa thẻ `<Override>` tương ứng trong `[Content_Types].xml`, dọn sạch các thẻ `<Relationship>` mồ côi trong `*.rels`, và tự động nâng cấp `ContentType` chính từ `macroEnabled.main+xml` về chuẩn OpenXML (`wordprocessingml.document.main+xml`, `spreadsheetml.sheet.main+xml`, `presentationml.presentation.main+xml`), ngăn chặn triệt để lỗi cảnh báo hỏng cấu trúc khi mở trên Microsoft Word/Excel.
  - **Edge Case Guard #2 (`ERR_SEC_007` — Precision Vietnamese PII Redactor with Financial Guard)**: Hàm `redact_pii_text` và chế độ `--redact-pii` kiểm soát chặt chẽ chống nhận diện nhầm số tiền lớn trong Báo cáo tài chính (BCTC):
    - *CCCD 12 số*: Bắt buộc khớp mã tỉnh hợp lệ (`001`–`096`) và mã thế kỷ/giới tính ở chữ số thứ 4 (`0`–`3`), đồng thời loại trừ các số đi kèm đơn vị tiền tệ/tỷ lệ (`VND`, `VNĐ`, `USD`, `EUR`, `đ`, `đồng`, `%`).
    - *Số điện thoại VN*: Khớp đúng đầu số nhà mạng di động Việt Nam (`+84`/`03x`, `05x`, `07x`, `08x`, `09x`) và không đứng trước đơn vị tiền tệ.
    - *Số tài khoản ngân hàng (STK)*: Bắt buộc có từ khóa ngữ cảnh neo phía trước (**Contextual Anchoring**: `STK`, `Số tài khoản`, `Tài khoản số`, `Account No`, `Acc No`, `A/C`) mới thực hiện che mờ, bảo vệ 100% các dãy số tiền 10–15 chữ số (`150000000000 VND`) và mã chứng từ nội bộ.
- **Phân Hệ 4 — Đường Ống Công Việc Khai Báo Đa Bước & Xử Lý Hàng Loạt (`core/pipeline/engine.py`)**:
  - `WorkflowPipeline`, `run_pipeline`, và `run_bulk_directory_pipeline`: Cho phép định nghĩa chuỗi tác vụ liên hoàn bằng JSON/YAML (ví dụ: `sanitize` ➔ `docx_apply_preset` ➔ `convert` ➔ `pdf_crypto_sign` ➔ `pdf_to_pdfa`).
  - **Edge Case Guard #4 (`ERR_PIPE_001` — Cloud 512MB RAM Worker Cap & Deterministic Temp Cleanup)**: Mặc định giới hạn luồng song song theo biến môi trường `max_workers = int(os.environ.get("MAX_JOB_WORKERS", 2))` để chống tràn bộ nhớ (OOM Killed) trên gói Render Cloud 512MB RAM, đồng thời thu hồi 100% các tệp trung gian `step_*_temp.*` bên trong khối `try...finally`.
- **Phân Hệ 5 — AST Code-to-Diagram Generator & Pure Draw.io Vector SVG Exporter (`core/diagrams/code_to_diagram.py`)**:
  - `scan_python_ast_to_spec`: Phân tích cây cú pháp trừu tượng Python (`ast`) để trích xuất tự động hệ thống Class, kế thừa, thuộc tính định kiểu và phương thức (`+`, `#`, `-`) sang PlantUML, Mermaid và Draw.io XML.
  - `scan_sql_or_prisma_to_erd_spec`: Phân tích câu lệnh SQL DDL (`CREATE TABLE`, `PRIMARY KEY`, `FOREIGN KEY REFERENCES`) và Prisma Schema sang sơ đồ thực thể liên kết (ERD).
  - `export_mxgraph_to_svg`: Kết xuất trực tiếp mô hình `<mxGraphModel>` sang định dạng ảnh vector `.svg` thuần Python không cần trình duyệt, bảo toàn lớp nền trắng chống đè chữ trên dây nối (`MX_INV_08`).
  - `generate_diagram_from_code`: Hỗ trợ tự động chèn sơ đồ kiến trúc vừa sinh vào cả tài liệu **Word (`.docx`)** lẫn **PowerPoint (`.pptx`)** qua tham số `inject_into`.
- **Phân Hệ 6 — Cầu Nối Giao Diện Desktop Native App Mode (`ui/webview_launcher.py`, `ui/desktop_app.py`, `main.py`)**:
  - **Edge Case Guard #5 (`ERR_UI_001` — Isolated `--user-data-dir` App Lifecycle)**: Hàm `build_app_mode_browser_command` và `launch_native_web_studio` truyền cờ `--user-data-dir=<temp_app_profile_dir>` cùng `--app=http://127.0.0.1:<port>` khi khởi chạy Microsoft Edge / Google Chrome. Cơ chế này tách biệt hoàn toàn tiến trình cửa sổ App khỏi trình duyệt đang chạy ngầm trên Windows, giúp `proc.wait()` giữ sống daemon server Python đúng đến khoảnh khắc người dùng bấm nút `X` đóng cửa sổ App, sau đó tự động tắt server và dọn sạch thư mục profile tạm trong khối `finally`.
  - Nâng cấp `ui/desktop_app.py` với nút bấm **🚀 Mở Full Web Studio (App Mode)** trên thanh Header và **Tab 4: 🛡️ Bảo Mật CDR, Excel-to-Brief & PDF/A**.
- **Mở Rộng Toàn Diện CLI, MCP Server (28 Tools) & REST API**:
  - Bổ sung 5 subcommands CLI mới (`xlsx-to-brief`, `xlsx-consolidate`, `run-pipeline`, `code-to-diagram`, `desktop-web`) và mở rộng `pdf-tools` (`crypto-sign`, `pdfa`), `doc-security` (`--deep-cdr`, `--redact-pii`).
  - Nâng tổng số công cụ MCP JSON-RPC 2.0 từ 23 lên **28 tools** (`pdf_crypto_sign_and_pdfa`, `xlsx_to_brief_report`, `xlsx_consolidate_workbooks`, `doc_run_pipeline`, `diagram_from_code`).

### Phiên Bản 4.5.0 (Cloudflare Defensive Security Audit & Document Sanitizer Suite)
- **Tích Hợp Kép Chuẩn Cloudflare Security Audit Harness (Dual-Level Integration)**:
  - **Cấp Độ 1 (Agentic Skill & Rule Level - `.agents/`)**:
    - **Skill `.agents/skills/security-audit/`**: Triển khai đầy đủ quy trình 6 giai đoạn của Cloudflare (`RECONNAISSANCE.md`, `HUNTING.md`, `VALIDATION-AND-REPORTING.md`, `ATTACK-CLASSES.md`, `OFFICE-AND-DOCUMENT-PARSERS.md`).
    - **Mô hình Phản Biện Đối Nghịch (Adversarial Verifier Sub-Agent)**: Phân định kết quả kiểm tra nghiêm ngặt thành 3 trạng thái:
      - `confirmed`: Lỗ hổng có vết taint trace hoàn chỉnh từ nguồn input đến sink thực thi, đã vượt qua thử thách phản biện.
      - `needs_validation`: Nghi vấn khả nghi cần kiểm chứng thêm ở runtime (không gán nhầm mức độ nghiêm trọng bừa bãi).
      - `rejected`: Mẫu kiểm tra đã được chứng minh vô hại kèm lý do phản biện cụ thể (`disprove_rationale`).
    - **Bộ Validator Node.js Thuần (Zero-Dependency)**: `validate-findings.cjs` và `validate-coverage-ledger.cjs` chạy cục bộ không cần npm package bên ngoài.
    - **Quy tắc `.agents/rules/rule_security_audit_standards.md`**: Thiết lập 5 bất biến bắt buộc `ERR_SEC_001`–`005` và nguyên tắc sandbox read-only.
  - **Cấp Độ 2 (Core Application & Feature Level - `core/security/`)**:
    - **`DocumentSecurityAuditor` (`core/security/document_auditor.py`)**: Động cơ kiểm tra tĩnh an ninh tài liệu phổ quát hỗ trợ DOCX, XLSX, PPTX, PDF, CSV, XML, DRAWIO, ZIP.
      - `ERR_SEC_001 (Macro & Malicious Code Execution)`: Quét `vbaProject.bin`, `word/vbaData.xml`, XLM macros, và PDF `/Launch`, `/JavaScript`.
      - `ERR_SEC_002 (Formula / DDE Injection)`: Dò tìm mọi ô bắt đầu bằng `=`, `+`, `-`, `@` gọi các hàm thực thi shell/mạng (`cmd|`, `powershell|`, `mshta|`, `cscript|`, `curl|`, `calc|`, `WEBSERVICE`, `HYPERLINK`).
      - `ERR_SEC_003 (Dangerous OLE Packages)`: Kiểm tra nhị phân OLE objects dưới `embeddings/oleObject*.bin`, bóc tách 256 bytes đầu phát hiện chữ ký Windows Executable (`MZ` PE header).
      - `ERR_SEC_004 (Hidden PII & Sensitive Metadata)`: Phát hiện tác giả `dc:creator`, `cp:lastModifiedBy`, ghi chú nội bộ `word/comments.xml`, và các sheet ẩn giấu cấp độ hệ thống `state="veryHidden"`.
      - `ERR_SEC_005 (Zip Slip Traversal & XXE)`: Dò tìm đường dẫn thoát thư mục `../` trong tên tệp ZIP và các khai báo thực thể bên ngoài `<!ENTITY` / `SYSTEM` trong cấu trúc XML.
    - **`DocumentSanitizer` (`core/security/sanitizer.py`)**: Động cơ tẩy rửa an toàn không phá hủy dữ liệu (Non-destructive Defuser):
      - *Cơ chế Vô Hiệu Hóa Formula Injection*: Thêm tiền tố nháy đơn `'` vào đầu công thức (`'=cmd|' /C calc'!A1`), biến công thức nguy hiểm thành chuỗi ký tự hiển thị vô hại mà không xóa dữ liệu người dùng.
      - *Tẩy Rửa Metadata Tác Giả*: Chuẩn hóa thông tin `dc:creator` và `cp:lastModifiedBy` trong `docProps/core.xml` thành định danh chung của hệ thống.
      - *Phục Hồi Worksheet Ẩn*: Xóa thuộc tính `state="veryHidden"` trong `xl/workbook.xml`, đưa toàn bộ bảng tính về trạng thái hiển thị bình thường.
  - **Tích Hợp Cổng 13 (Gate 13: Security & Sanitizer Shield) Vào `tools/unified_qa_diagnostic.py`**:
    - Bổ sung `DocumentSecurityDiagnosticEngine` vào hệ thống chẩn đoán thống nhất.
    - Hỗ trợ các cờ dòng lệnh `--security` và `--sanitize` trong lệnh `python ai_tools_cli.py qa-audit`.
  - **Hợp Nhất CLI Subcommand Mới (`doc-security`)**:
    - Đăng ký lệnh `python ai_tools_cli.py doc-security <doc> [--sanitize] [-o output] [--format {terminal,json}]`.
  - **Cung Cấp REST API Endpoints Chuyên Biệt**:
    - `POST /api/security/scan`: Quét an ninh tĩnh trả về báo cáo JSON 3 trạng thái.
    - `POST /api/security/sanitize`: Tẩy rửa an toàn và trả về đường dẫn tải tệp sạch.
    - Hỗ trợ tham số `security: true` trong endpoint `/api/qa/audit`.
  - **Đăng Ký Model Context Protocol (MCP) Tool Thứ 23**:
    - Bổ sung công cụ `doc_security_audit` vào `server/mcp_server.py` và facade tại root `mcp_server.py`.
  - **Nâng Cấp Giao Diện Web Studio (Tab QA Panel)**:
    - Bổ sung công tắc `[x] Kiểm định an ninh tài liệu & Tẩy rửa mã độc (Gate 13: Cloudflare Security & PII Shield)`.
    - Hiển thị huy hiệu an ninh và danh sách chi tiết các nguy cơ an ninh phát hiện được.
  - **Hệ Thống Kiểm Thử Tự Động (305 Tests, 100% Pass Rate)**:
    - Bổ sung `tests/test_document_security.py` (15 test cases toàn diện) và mở rộng `tests/test_mcp_server.py` lên 23 tools.

### Phiên Bản 4.4.0 (Digital Yearbook & Pre-Press Photobook Studio Suite)
- **Phân Hệ Digital Yearbook & Pre-Press Photobook Studio (`core/yearbook/`)**:
  - **Trình Xem 3D Flipbook Trực Quan Trên Web Studio (Tab 6)**:
    - Hiển thị 2 trang đôi mở phẳng (Layflat Photobook) với rãnh gáy sách mô phỏng bóng đổ chân thực (Realistic Center Spine Gutter Shadow).
    - Tương tác kéo thả hoàn chỉnh (HTML5 Drag & Drop): Kéo ảnh từ Khay ảnh (Photo Tray) vào khung (Slot), hoặc kéo ảnh giữa 2 khung để hoán đổi vị trí (Slot-to-Slot Swapping) tức thì.
    - Sửa trực tiếp lời tựa (Inline Live Editable Caption): Bấm chuột vào dòng chữ dưới chân trang hoặc trên khung polaroid để sửa văn bản, tự động đồng bộ hai chiều với bảng điều khiển Inspector.
    - Công tắc Vạch Xén In Ấn (Pre-Press Guides Overlay): Bật/tắt 3 đường biên kỹ thuật: Lề xén 3mm (Bleed - Đỏ), Vạch cắt thành phẩm (Trim line - Xanh dương), và Vùng an toàn (Safe zone - Xanh lá).
    - Khởi tạo Zero-Config Mock Demo Album: Ngay khi mở giao diện, album mẫu 8 trang đôi cử nhân ("KỶ YẾU TỐT NGHIỆP 2026 - BIỆN TRUNG TÍN") được tự động nạp sẵn với hình ảnh vector mỹ thuật cao.
  - **9 Mẫu Layout Chuẩn Tạp Chí (Editorial Layout Templates)**:
    - `cover_hero` (1 ảnh): Bìa Kỷ Yếu Cử Nhân với Typography hoàng gia, niên khóa và lời đề tựa.
    - `editorial_quote` (1 ảnh): Trang Mở Đầu với đoạn văn cảm nghĩ, tri ân gia đình, thầy cô và bạn bè.
    - `symmetric_duo` (2 ảnh): Song Hành Cân Xứng cho ảnh chụp đôi cùng bạn thân hoặc cha mẹ.
    - `magazine_grid_4` (4 ảnh): Tạp chí 4 khung gồm 1 ảnh Hero lớn dẫn hướng thị giác và 3 ảnh vệ tinh.
    - `polaroid_scrapbook` (3 ảnh): Khung ảnh Polaroid viền trắng đổ bóng giấy, góc nghiêng tự nhiên kèm ghi chú lưu bút.
    - `mosaic_gallery_6` (6 ảnh): Phòng Trưng Bày Mosaic đa sắc, cân đối giữa ảnh ngang 16:9 và ảnh vuông 1:1.
    - `classroom_grid_8` (8 ảnh): Lưới Kỷ Yếu Tập Thể 2x4 chuẩn quốc tế cho nhóm đồ án và bạn bè.
    - `megagrid_12` (12 ảnh): Đại Lưới 12 ô gom trọn 50–100+ ảnh selfie hậu trường và khoảnh khắc vui nhộn.
    - `panoramic_spread` (1 ảnh): Ảnh Ngang Toàn Cảnh trải rộng xuyên suốt cả 2 trang mở phẳng không đứt gãy.
  - **Bộ 4 Bất Biến Kỹ Thuật Công Nghiệp (Technical Invariants)**:
    - *ERR_PDF_FONT_001 (Vietnamese Unicode Font Embedding)*: Tự động dò tìm và nhúng font hệ thống TTF Unicode (Arial, Times New Roman, LiberationSans, DejaVuSans), khử hoàn toàn lỗi rụng dấu tiếng Việt trên PyMuPDF.
    - *ERR_EXIF_001 (EXIF Orientation Normalization)*: Chuẩn hóa góc quay hình ảnh máy ảnh và điện thoại qua `ImageOps.exif_transpose` trước khi xử lý, chống hiện tượng ảnh chụp dọc bị lật ngang.
    - *ERR_IMG_CROP_001 (Top-Biased 35% Aspect Crop)*: Thuật toán cắt ảnh theo tỷ lệ khung hình với cơ chế `top_bias=0.35` bảo toàn 100% gương mặt và mũ tốt nghiệp, chống viền trắng letterbox.
    - *ERR_RAM_001 (Streamed Disk Packaging)*: Cơ chế stream tệp trực tiếp từ đĩa vào file ZIP và dọn rác bộ nhớ định kỳ sau mỗi 4 trang đôi, duy trì mức tiêu thụ RAM $\le 50\text{MB}$ khi xử lý 50–100+ ảnh gốc dung lượng lớn.
  - **Gói Bàn Giao Thợ Thiết Kế 5-trong-1 (Designer Handoff Bundle .zip)**:
    - `01_Digital_Yearbook_Spread_Preview.pdf`: Bản PDF mở phẳng 2 trang đôi duyệt tổng thể bố cục.
    - `02_Print_Ready_Pages_300DPI.pdf`: Bản PDF in ấn từng trang đơn chuẩn 300 DPI kèm Bleed 3mm và 4 cặp Crop Marks.
    - `02_Original_Photos_Sorted/`: Thư mục ảnh gốc độ phân giải cao được phân loại và đổi tên có hệ thống (`Page_XX_Slot_YY.jpg`).
    - `03_Design_Brief_And_Captions.docx`: Bản đặc tả thiết kế Word chuẩn OpenXML chứa bảng phân bổ chi tiết từng trang, tuân thủ tuyệt đối các bất biến `<w:cantSplit/>`, `<w:tblHeader/>`, `<w:vAlign w:val="center"/>` và The Last Paragraph Rule.
    - `03_Designer_Handoff_Brief.md` & `04_Yearbook_Project_Spec.json`: Bản tóm tắt Markdown và hồ sơ đặc tả JSON cho thợ thiết kế.
  - **Tích Hợp Toàn Diện Server, CLI & MCP**:
    - REST API: 5 endpoints mới (`/api/yearbook/mock-album`, `/api/yearbook/presets`, `/api/yearbook/auto-layout`, `/api/yearbook/render-preview`, `/api/yearbook/export`) và background job runner.
    - CLI Subcommands: `yearbook-layout` (tự động dàn trang từ ảnh hoặc cờ `--mock`) và `yearbook-build` (đóng gói ZIP handoff).
    - Model Context Protocol (MCP): 2 công cụ mới `yearbook_auto_layout` và `yearbook_export` nâng tổng số MCP tools lên 22 công cụ.
    - Kiểm thử: 15 bài test tự động chuyên sâu nâng tổng quy mô kiểm thử toàn hệ thống lên 289 tests (100% Passed).

### Phiên Bản 4.3.0 (Excel Formula Playbook & Adaptive Matrix Enterprise Suite)

- **Động Cơ Quản Trị & Tự Động Hóa Bảng Tính Cấp Doanh Nghiệp (Enterprise Excel Data Suite)**:
  - **Khắc Phục Triệt Để 4 Nhóm Bẫy Lỗi Biên (Edge Cases)**:
    - *ERR_XLSX_009 (15-Digit Precision & Accounting Negative Guard)*: Ép kiểu chuỗi `@` cho số thẻ/CCCD/STK $\ge 12$ số; tự động phân tích ngày tháng tiếng Việt (`26/05/2026`, `ngày 26 tháng 05 năm 2026`) thành `datetime.date` với format `DD/MM/YYYY`; chuyển số âm kế toán `(100.000)` thành float âm `-100000` và `-` thành `0` kèm định dạng kế toán; chuẩn hóa Locale thập phân Việt Nam `20.000,50` $\rightarrow$ `20000.50`.
    - *ERR_XLSX_010 (XML Illegal Character Sanitization & 32k Limit)*: Quét sạch ký tự điều khiển ASCII `[\x00-\x08\x0b\x0c\x0e-\x1f]` chống lỗi tệp OpenXML bị hỏng ("The file is corrupt"); tự động cắt tỉa chuỗi dài tối đa 32.000 ký tự (ngưỡng an toàn dưới 32.767) kèm hậu tố `... [Xem tiếp tài liệu gốc]`.
    - *ERR_XLSX_011 (_xlfn Modern Formula Normalization)*: Tự động bổ sung tiền tố `_xlfn.` cho tất cả các hàm mảng động và hiện đại (`XLOOKUP`, `CONCAT`, `TEXTJOIN`, `UNIQUE`, `FILTER`, `SORT`, `SORTBY`, v.v.) chống lỗi `#NAME?` khi mở trên Excel 2016/2019 cũ.
    - *ERR_XLSX_012 (Gridlines, AutoFilter & Unlocked Row Height)*: Luôn bảo đảm bật đường lưới `ws.views.sheetView[0].showGridLines = True`; tự động gắn vùng lọc `ws.auto_filter.ref`; để `ws.row_dimensions[r].height = None` cho dòng dữ liệu cho phép Excel tự co giãn theo dòng `\n` tránh cắt cụt văn bản.
  - **6 Phân Hệ Nghiệp Vụ Doanh Nghiệp Trọng Yếu**:
    - *Phân hệ 1 (Automated Conditional Formatting)*: Highlight trạng thái tự động theo bảng màu pastel doanh nghiệp (Đỏ nhạt `#FEE2E2` cho "Chưa duyệt"/"Từ chối", Xanh lá `#DCFCE7` cho "Đã duyệt"/"Hoàn thành", Vàng nhạt `#FEF9C3` cho "Đang chờ"); nhúng Data Bars trực tiếp trong ô số liệu cho tỷ lệ % / KPI; dải màu nhiệt Color Scales 3 điểm.
    - *Phân hệ 2 (Data Validation & Dropdown Lists)*: Gắn danh mục chọn sẵn (Dropdown Lists) chống gõ sai chính tả; kiểm soát độ dài ký tự (CCCD bắt buộc 12 số, MST 10 số); giới hạn khoảng ngày nhập liệu năm 2026.
    - *Phân hệ 3 (Native In-Sheet Charts - ERR_DIAG_003)*: Nhúng biểu đồ OpenXML sống (`BarChart`, `LineChart`, `PieChart`) trực tiếp vào worksheet; tự động neo an toàn dưới bảng dữ liệu hoặc bên cạnh bảng, tuyệt đối không đè lên số liệu.
    - *Phân hệ 4 (Selective Cell Protection & Formula Locking - ERR_XLSX_013)*: Mọi ô chứa công thức và tiêu đề mặc định khóa `locked=True`; chỉ mở khóa vùng nhập liệu được chỉ định; kích hoạt `ws.protection.enable()` bảo vệ bảng tính.
    - *Phân hệ 5 (Big Data Streaming Mode - ERR_XLSX_014)*: Module `core.xlsx.streaming` hỗ trợ `stream_read_xlsx` (`read_only=True`) và `stream_write_xlsx` (`write_only=True`), duy trì RAM $< 50\text{MB}$ khi xử lý tệp $> 10\text{MB}$ hoặc $> 30.000$ dòng, giải quyết triệt để rủi ro OOM trên gói Free Render.com 512MB.
    - *Phân hệ 6 (Executive Index Dashboard & 2-Way Hyperlink Navigation)*: Tự động sinh trang Mục Lục đầu tiên với banner Navy `#0F294A`, bảng tổng quan thống kê số dòng và công thức hyperlink `=HYPERLINK(...)` nhảy trang nhanh; chèn nút `[⬅ Về Trang Mục Lục]` tại `A1` của các sheet con khi chuyển đổi tài liệu nhiều bảng.
- **Hệ Thống Cẩm Nang Công Thức Excel Tinh Hoa & Cheat Sheet System (`core.xlsx.formula_playbook`)**:
  - **Catalog 32+ Hàm Excel Thực Chiến Theo 4 Chiều Toàn Diện**:
    - *Tên hàm & Phân loại*: 6 nhóm nghiệp vụ lớn (Tra cứu `lookup`, Tổng hợp `aggregation`, Chuỗi văn bản `text`, Điều kiện `logic`, Mảng động `dynamic`, Tài chính & Thời gian `financial`).
    - *Cú pháp & Công thức mẫu*: Công thức chuẩn cú pháp kèm ví dụ thực tế tham chiếu ô tính bắt đầu bằng dấu `=`, không dùng placeholder giả.
    - *Dùng trong trường hợp nào*: Tình huống kế toán, tài chính, kiểm toán, và xử lý dữ liệu thực tế (báo cáo doanh thu, trích xuất dữ liệu, kiểm tra trùng lặp).
    - *Tại sao lại chọn hàm đó của Excel*: Phân tích kiến trúc kỹ thuật sâu sắc, so sánh trực diện ưu điểm vượt trội so với các hàm cổ điển (như `XLOOKUP` vs `VLOOKUP`, `IFS` vs lồng `IF`, `TEXTJOIN` vs `CONCATENATE`, `SUMIFS` vs `SUMPRODUCT`).
  - **Động Cơ Xuất Sổ Tính Excel 2 Sheets (`generate_formula_playbook_xlsx`)**:
    - *Sheet 1 ('Formula_Playbook')*: Bảng tra cứu chuẩn Enterprise Navy (`0F294A`), phông Consolas cho công thức, tự động tính chiều cao dòng theo độ dài văn bản (`ERR_XLSX_005`), cố định dòng tiêu đề `freeze_panes = 'A4'`, thiết lập khổ in A4 ngang Print-Ready (`ERR_XLSX_008`).
    - *Sheet 2 ('Live_Interactive_Formulas')*: Bảng dữ liệu nhân viên thực tế đang chạy sống động các công thức KPI: `=SUMIFS(...)`, `=COUNTIFS(...)`, `=AVERAGEIFS(...)`, `=XLOOKUP(...)`, `=IFS(...)`, `=TEXTJOIN(...)` giúp người dùng mở tệp là thấy ngay kết quả tính toán.
  - **Tích Hợp Toàn Bộ Giao Thức (CLI / REST / MCP / Web Studio)**:
    - *CLI*: `python ai_tools_cli.py xlsx-formulas [--category <cat>] [--search <query>] [--export <file.xlsx>] [--json]`.
    - *REST API*: `GET /api/xlsx/formulas` (hỗ trợ lọc danh mục & tìm kiếm) và `GET /api/xlsx/formulas/export` (tải tệp Excel 2 sheets trực tiếp).
    - *MCP Tool #20*: `xlsx_formula_helper` hỗ trợ AI Agent tra cứu nhanh cú pháp, bài toán thực tế, và lý do lựa chọn hàm.
    - *Web Studio UI*: Sub-tab 4 `Formula Playbook & Live Demo` trong `Spreadsheet Studio` với giao diện thẻ kính Glassmorphism, live search debounced, category pills, 1-click copy công thức vào clipboard, và nút xuất Excel 2 sheets.
- **Atomic Office Tools Suite & Web Quick Patch (MCP / CLI / Web Studio)**:
  - **8 Công Cụ Văn Phòng Hạt Nhân Cho AI Agent (MCP JSON-RPC 2.0)**: Mở rộng Tool Catalog từ 11 lên 19 công cụ chuẩn:
    - 4 Excel Tools: `xlsx_set_cell`, `xlsx_set_range`, `xlsx_insert_rows`, `xlsx_find_anchor`.
    - 4 Word Tools: `docx_inspect_placeholders`, `docx_replace_text`, `docx_insert_paragraph`, `docx_insert_table`.
    - Bảo toàn nghiêm ngặt các Invariants: `ERR_XLSX_004` (Safe Merged-Cell Handling: tự động tìm Top-Left của vùng merged chống crash read-only), `ERR_XLSX_007` (Identifier Leading Zero: bảo toàn chuỗi `'0912345678'` với `@`), `ERR_DOCX_002` (Run Text Overwrite: bảo toàn kiểu chữ, phông, màu sắc của run OpenXML), `ERR_DOCX_003` (OpenXML table invariants `<w:cantSplit/>`, `<w:tblHeader/>`).
    - Chuẩn hóa hai chiều token keys: hỗ trợ cả `KEY` và `{{KEY}}` linh hoạt.
  - **4 Lệnh Con CLI Mới Trong `cli/ai_tools.py`**:
    - `python ai_tools_cli.py xlsx-set <file.xlsx> --sheet <sheet> --cell <cell> --value <val> [--formula <f>]`
    - `python ai_tools_cli.py xlsx-insert-rows <file.xlsx> --sheet <sheet> --start-row <row> [--count <n>] [--proto <row>]`
    - `python ai_tools_cli.py docx-tokens <template.docx> [--json]`
    - `python ai_tools_cli.py docx-replace <doc.docx> --data <patch.json>`
  - **Giao Diện DOCX Quick Patch & Two-Way Sync Trên Web Studio**:
    - Sub-tab mới `DOCX Quick Patch` trong `panel-documents` với bộ chuyển đổi 2 chế độ (Form Trực Quan $\leftrightarrow$ JSON Editor) đồng bộ dữ liệu thời gian thực.
    - Tự động gọi endpoint `/api/docx/inspect-tokens` để bóc tách cây AST biến `{{placeholders}}` ngay khi tải tệp Word lên.
    - Phân hệ Spreadsheet: Nút preset mẫu `Sửa Nhanh Ô Tính (Cell Patch)` và tự động nhận diện danh sách trang tính (`sheet_names`).
  - **Gia Cố File Watcher (`vite.config.js`)**: Bổ sung mẫu ignore `scratch/**`, `**/*.tmp`, `**/~$*`, `**/.~lock.*` triệt tiêu lỗi EBUSY crash Node.js khi thao tác Office.
- **Kéo-Thả Đổi Thứ Tự Trang Trực Quan & Chèn Trang Trắng A4 (Drag-and-Drop Page Reordering)**:
  - `core/converter/pdf_toolkit.py` (`reorganize_pdf_pages`): Hỗ trợ cờ `is_blank=True` cho phép tạo trang trắng chuẩn A4 ($595.32 \times 841.92\text{ pt}$) tại bất kỳ vị trí nào trong tài liệu.
  - Web Studio: Tích hợp HTML5 Drag-and-Drop (`draggable="true"`) cho phép nắm chuột kéo thả đổi chỗ các thẻ thumbnail 120 DPI; bổ sung nút `+ Chèn Trang Trắng` trên thanh công cụ Organizer.
- **Chuyển Đổi Hai Chiều Ảnh ↔ PDF (Images to PDF & 300 DPI Export ZIP)**:
  - `convert_images_to_pdf()`: Hợp nhất nhiều ảnh (.jpg, .jpeg, .png, .webp) thành tệp PDF chuẩn in ấn, tự động chuẩn hóa góc quay EXIF (`ImageOps.exif_transpose`), tự động tính toán margin và định hướng trang (Auto/Portrait/Landscape).
  - `export_pdf_to_images_zip()`: Kết xuất toàn bộ các trang PDF thành ảnh PNG/JPG độ phân giải cao 300 DPI đóng gói trong tệp `.zip` nén chuẩn deflate.
- **Ký Tên Điện Tử Trực Tiếp Trên Web Studio & Đóng Dấu Bù Góc Xoay (Sign & Annotate PDF)**:
  - **Giao diện Web Studio**: Tích hợp nút tác vụ `Ký Tên & Con Dấu` trong phân hệ `PDF Pro Studio` (`Document Hub`), cung cấp bảng điều khiển chuyên nghiệp:
    - *Interactive Canvas Signature Pad*: Vẽ tay mượt mà (chuột, bút cảm ứng, ngón tay), hỗ trợ 3 màu mực chuẩn (`#0033aa` xanh doanh nghiệp, `#111827` đen pháp lý, `#dc2626` đỏ mộc), nút xóa nét tức thì.
    - *Tải lên ảnh mộc dấu / chữ ký*: Hỗ trợ PNG/JPG tách nền trong suốt.
    - *Phạm vi trang linh hoạt*: Chọn trang cuối (chuẩn hợp đồng), trang đầu, tất cả các trang, hoặc danh sách trang tùy biến (`1, 3, 5` hoặc `2-4`).
    - *Preset vị trí & thanh trượt tỉ lệ*: 4 góc chuẩn (`Góc Dưới Phải`, `Góc Dưới Trái`, `Giữa Trang`, `Góc Trên Phải`) và thanh trượt co giãn kích thước từ 12% đến 45%.
    - *In kèm chú thích người ký*: Tự động chèn 3 dòng vector Unicode sắc nét bên dưới chữ ký (Chức danh, Họ tên, Ngày ký).
  - `sign_pdf_with_stamp()`: Hỗ trợ cả base64 và file path, bù góc xoay tự động qua `derotation_matrix`, vector text metadata injection.
  - Thuật toán bù góc xoay: Tự động phát hiện trang có góc quay 90°, 180°, 270° (`page.rotation != 0`), áp dụng ma trận biến đổi ngược `v_rect * page.derotation_matrix` kết hợp cờ `rotate=page.rotation` chống lệch tọa độ chữ ký.
- **Làm Phẳng Tài Liệu Hai Chế Độ & Triệt Tiêu Dấu Vết Metadata (Flatten & Sanitize PDF)**:
  - `flatten_and_sanitize_pdf()`:
    - **Vector Bake Mode** (`doc.bake(annots=True, widgets=True)`): Khóa cứng biểu mẫu AcroForm, chữ ký và chú thích vào luồng nội dung vector, giữ nguyên 100% độ sắc nét của chữ.
    - **Strict Raster Flattening** (`page.get_pixmap(dpi=200)`): Chuyển hóa toàn bộ trang thành ảnh tĩnh bitmap, ngăn chặn tuyệt đối việc bóc tách hay trích xuất con dấu và chữ ký.
    - **Metadata Purge** (`doc.scrub()`): Tẩy sạch tác giả, phần mềm tạo file, tiêu đề, ngày chỉnh sửa, XML XMP và các tệp đính kèm ẩn.
- **Chuỗi Hành Động Liên Hoàn 1-Click (Workflow Chaining UX)**:
  - Hiển thị thanh thao tác nhanh dưới kết quả cho phép người dùng chuyển tiếp tệp vừa tạo sang các công cụ khác (Nén ➔ Khóa AES-256 ➔ Đóng dấu giáp lai ➔ Xuất ảnh ZIP) mà không cần tải lại file từ đầu.
- **Co Giãn Kích Thước Thẻ Slide Thông Minh (Adaptive PPTX Card Sizing)**:
  - `core/pptx/writer.py` & `core/pptx/docx_to_pptx.py`:
  - Giải quyết dứt điểm tình trạng thẻ slide ít chữ nhưng khung vẽ quá to gây mất đối xứng: tự động tính toán chiều cao thẻ `card_h` và dải phân bố dọc theo số dòng thực tế của nội dung (`est_lines`), co gọn khung và căn lề hài hòa.
- **Ma Trận Chuyển Đổi 5 Chiều: Trích Xuất Bảng Biểu Ngữ Nghĩa Sang Excel (.xlsx)**:
  - `core/converter/table_extractor.py` & root facade `table_extractor.py`: Trích xuất dữ liệu bảng biểu có cấu trúc từ PDF, DOCX (OpenXML) và Markdown sang workbook Excel đa sheet định dạng Enterprise.
  - **Bộ Hậu Xử Lý Ma Trận & Thu Gom Toạ Độ Cột (Matrix Post-Processor & Ghost Column Pruning)**:
    - *Ưu tiên chiến lược Vector Lines Strict*: Thiết lập `page.find_tables(strategy="lines_strict")` bám sát 100% các đường viền vector thực tế của bảng, triệt tiêu hoàn toàn hiện tượng phình to 23 cột ma và 10 cột rỗng xen kẽ do suy diễn khoảng hở văn bản (fallback `page.find_tables()` linh hoạt nếu bảng không có viền kẻ).
    - *Thuật toán cắt bỏ cột ma (`_prune_ghost_columns`)*: Quét toàn bộ ma trận dữ liệu và tự động loại bỏ các cột rỗng ảo $100\%$ giá trị `None`/khoảng trắng, ép dữ liệu về đúng 9 cột dữ liệu thực tế.
    - *Ghép nối bảng đa trang liền mạch (`_stitch_multipage_tables`)*: Nhận diện bảng nối tiếp qua nhiều trang dựa trên số cột đối xứng và chữ ký tiêu đề (header signature matching), tự động khử 2 dòng tiêu đề lặp lại ở các trang tiếp theo, gom toàn bộ 106 dòng qua 7 trang vào đúng **1 Sheet duy nhất** (`Bieu_Phi_Dich_Vu`) thay vì chia cắt thành nhiều sheet rời rạc.
    - *Tái tạo cấu trúc ô gộp tiêu đề 2 tầng (`_detect_table_header_structure`)*: Tự động phân tích quan hệ nhãn cha - con, tái tạo chính xác 5 dải ô gộp chuẩn: `A1:A2`, `B1:B2`, `C1:C2` (gộp dọc) và `D1:F1`, `G1:I1` (gộp ngang) với màu nền Navy `#0F294A` và chữ trắng in đậm.
    - *Khóa cố định dòng tiêu đề (`freeze_panes`)*: Tự động thiết lập `ws.freeze_panes = 'A3'` cho bảng tiêu đề 2 tầng và `ws.freeze_panes = 'A2'` cho bảng tiêu đề 1 tầng trên toàn bộ các định dạng PDF, DOCX, Markdown, tối ưu trải nghiệm cuộn bảng tính hàng trăm dòng.
  - Lọc bảng hành chính Nghị định 30 (`_is_layout_header_or_signature_table`): Tự động phát hiện và bỏ qua bảng căn lề 2 cột chứa Quốc hiệu, Tiêu ngữ, Nơi nhận và Chữ ký.
  - Xử lý an toàn ô gộp `gridSpan` (colspan) & `vMerge` (rowspan) qua thẻ XML `<w:tc>` con: Định dạng viền/nền vùng gộp mà không gán lặp thuộc tính `.value`, triệt tiêu lỗi `AttributeError: 'MergedCell' object attribute 'value' is read-only` (`ERR_XLSX_004`).
  - Chuẩn hóa tên Sheet Excel (`_sanitize_sheet_title`, `_derive_pdf_sheet_title`): Khử 7 ký tự cấm `\/?*:[]`, giới hạn $\le 28$ ký tự, chống trùng lặp case-insensitively, tự động đặt tên trang tính trang nhã theo ngữ cảnh tài liệu (`Bieu_Phi_Dich_Vu`).
  - Bảo tồn chuỗi mã định danh (`ERR_XLSX_007`): Giữ nguyên số 0 ở đầu cho mã CCCD, mã số thuế, số tài khoản bằng định dạng chuỗi `@`.
  - Nối thông Routes 10, 11, 12 trong `core/converter/engine.py` (PDF $\to$ XLSX, DOCX $\to$ XLSX, MD $\to$ XLSX) và bổ sung wrapper tiện ích trong `core/converter/pdf_toolkit.py` (`extract_pdf_tables_to_xlsx`).
- **Bộ Tổng Hợp Slide Cấp Quản Lý Cho Tài Liệu Dài $\ge 100$ Trang (Executive Deck Synthesizer)**:
  - `core/pptx/docx_to_pptx.py`: Cơ chế Target Deck Sizing (`auto`, `executive`, `standard`, `comprehensive`). Tự động kích hoạt chế độ `executive` khi tài liệu $> 15.000$ từ hoặc $> 40$ đề mục.
  - Semantic Fluff Pruning (`RE_FLUFF_PARAGRAPH`): Loại bỏ câu văn đệm, rườm rà mang tính hình thức.
  - Quantitative KPI Prioritization (`RE_QUANTITATIVE_SIGNAL`): Ưu tiên bảo tồn câu có số liệu định lượng (%, tỷ đồng, USD, thời gian, đơn vị đo).
  - Thắt chặt độ sâu đề mục: Gộp các bullet H3/H4 technical minutiae vào H2 cha thay vì sinh slide rời.
  - Bảo toàn 100% bảng biểu (`table`) và sơ đồ kiến trúc (`diagram`/`ascii_diagram`).
  - Khống chế trần cứng $\le 25$ slides cho bài thuyết trình cấp điều hành.
  - Web Studio: Tải ảnh xem trước 1080p theo lô (Batch Progressive Loading: 16 slides/lần).
- **Mục Lục Tương Tác Đa Định Dạng 3 Chiều (Interactive Multi-Format Table of Contents)**:
  - `core/converter/markdown.py` & root facade `markdown_converter.py`:
    - **Click Nhảy Trực Tiếp (1-Click Direct Jump)**: Cho phép nhấp chuột vào bất kỳ mục nào trên mục lục để nhảy ngay lập tức tới dòng/tiêu đề tương ứng trên cả 3 định dạng: Markdown (`.md`), Word (`.docx`), và PDF (`.pdf`).
    - **Markdown Slug Chuẩn GFM**: Tự động chuyển đổi tiêu đề thành slug tương thích GitHub Flavored Markdown và trình đọc VS Code (`slugify_markdown_heading`), loại bỏ dấu tiếng Việt và ký tự đặc biệt, nối từ bằng dấu gạch ngang đơn (`-`).
    - **Word Bookmark & Hyperlink Engine**: Tự động sinh thẻ định danh OpenXML `<w:bookmarkStart>` và `<w:bookmarkEnd>` cho toàn bộ các đề mục (giới hạn chuẩn $\le 40$ ký tự ASCII không dấu theo chuẩn ISO/IEC 29500), đồng thời bọc dòng mục lục trong `<w:hyperlink w:anchor="...">`.
    - **Clean Publication Style**: Triệt tiêu màu xanh gạch chân lòe loẹt của hyperlink mặc định; giữ trọn vẹn màu sắc, phông chữ và phong cách xuất bản chuyên nghiệp của tài liệu.
    - **Bảo Toàn Liên Kết Nội Bộ Trên PDF (`LINK_GOTO`)**: Khi xuất DOCX sang PDF bằng LibreOffice hoặc MS Word COM, toàn bộ các liên kết mục lục tự động chuyển hóa thành PDF internal link annotations với tọa độ điểm nhảy chính xác.
    - **Pass 1 Pre-scan & Fuzzy Heading Catalog**: Quét trước danh mục đề mục 2 chiều (slug, clean title, line index) cho phép liên kết chính xác cả với mục lục thủ công và mục lục sinh tự động.
- **Mở Rộng Giao Thức CLI & REST API**:
  - CLI: Mở rộng lệnh `python ai_tools_cli.py pdf-tools` với 5 hành động mới: `images-to-pdf`, `export-images-zip`, `sign`, `flatten`, `reorganize`; bổ sung cờ `--scope {auto,executive,standard,comprehensive}` vào `docx-to-pptx`, `pdf-to-pptx`; bổ sung lệnh `python ai_tools_cli.py md-toc <file.md> [--update]` để xem trước và tự động đồng bộ mục lục tương tác trong tệp Markdown.
  - REST API: Bổ sung 4 endpoint chuyên biệt trong `server/app.py`: `POST /api/pdf/images-to-pdf`, `POST /api/pdf/export-images-zip`, `POST /api/pdf/sign`, `POST /api/pdf/sanitize`; hỗ trợ dải slide `1-16`, `17-32` tại `/api/pptx/preview`.
- **Bộ Kiểm Thử Chất Lượng (Quality Gate Suite Expansion)**:
  - Bổ sung `tests/test_pdf_ilovepdf_features.py` (5 tests).
  - Bổ sung `tests/test_universal_table_xlsx.py` (6 tests).
  - Bổ sung `tests/test_executive_deck.py` (6 tests).
  - Bổ sung `tests/test_interactive_toc.py` (5 tests).
  - Bổ sung `tests/test_pdf_table_matrix_stitching.py` (7 tests).
  - Toàn bộ test suite tự động vượt qua 100% (211 tests, 100% Pass Rate).

### Phiên Bản 4.1.0 (PDF Pro Studio & Visual Page Organizer Enterprise Suite)
- **Nhóm Bảo Mật & Tuân Thủ Pháp Lý (True Redaction & AES-256 Encryption)**:
  - `core/converter/pdf_toolkit.py` & root facade `pdf_toolkit.py`:
    - `redact_sensitive_pdf()`: Áp dụng `page.add_redact_annot()` kết hợp `page.apply_redactions(images=fitz.PDF_REDACT_IMAGE_PIXELS)` triệt tiêu vĩnh viễn byte văn bản trong AST và pixel ảnh raster bên dưới (**`ERR_PDF_001`**).
    - Bộ Regex tự động: Căn cước công dân (`cccd`), Số điện thoại VN (`phone`), Email (`email`), Thẻ tín dụng/ATM (`bank_card`), Mức lương & Thu nhập (`salary`), kèm từ khóa tùy chỉnh và tùy chọn màu che phủ (đen/trắng).
    - `protect_pdf()`: Mã hóa chuẩn quân đội AES-256 (`fitz.PDF_ENCRYPT_AES_256`), phân tách Mật khẩu xem (`user_password`) và Mật khẩu quản trị (`owner_password`), hỗ trợ phân quyền in ấn (`allow_print`), sao chép văn bản (`allow_copy`), và chỉnh sửa (`allow_modify`).
- **Nhóm Hành Chính & Kế Toán Việt Nam (Dấu Giáp Lai Đa Trang & Chuẩn Hóa A4 Bates)**:
  - `stamp_giap_lai_pdf()`: Tự động phân rã ảnh con dấu tròn/vuông thành $N$ lát cắt dọc tương ứng với $N$ trang văn bản, bổ sung độ chồng lấn 8% (**`ERR_PDF_002`**) chống khe hở khi in ấn, căn lề mép ngoài (mép phải chuẩn hợp đồng hoặc mép trái).
  - `normalize_and_number_pdf()`: Vector Embedding qua `show_pdf_page()` căn giữa chuẩn A4 ($595.28 \times 841.89\text{ pt}$) bảo toàn độ nét vector 100% (**`ERR_PDF_003`**), tự động nhận diện hướng trang (Auto/Portrait/Landscape), kết hợp đóng số pháp lý Bates định dạng Unicode tiếng Việt.
- **Nhóm Bóc Tách Bảng Biểu Trực Tiếp Sang Excel (`extract_pdf_tables_to_xlsx`)**:
  - Tận dụng thuật toán C++ `page.find_tables()` có sẵn trong PyMuPDF để phát hiện cấu trúc bảng biểu.
  - Sinh bảng tính `.xlsx` chuẩn doanh nghiệp với phong cách Navy Header `#0F294A`, dòng xen kẽ zebra, viền mảnh.
  - Tuân thủ nghiêm ngặt **`ERR_XLSX_007`** (bảo toàn số 0 ở đầu mã CCCD, Mã số thuế, Số tài khoản ở dạng chuỗi `@`) và **`ERR_XLSX_008`** (khổ in A4 ngang `fitToPage = True`).
- **Trình Biên Tập Trang Trực Quan Trên Web Studio (Visual Page Organizer)**:
  - `index.html`, `src/index.css`, `src/main.js`:
    - Tích hợp Sub-tab 6 **PDF Pro Studio** vào Document Hub.
    - Lưới xem trước Thumbnail 120 DPI (`/api/pdf/pages/preview`).
    - Nút xoay 90° tức thời `[↻]` với phản hồi hoạt họa CSS transform mượt mà.
    - Nút xóa trang `[🗑️]` với hiệu ứng mờ bán trong suốt, viền đỏ, và huy hiệu "ĐÃ XÓA".
    - Bộ 6 thẻ thao tác chuyển đổi tức thời qua Action Pills.
- **Mở Rộng Giao Thức (REST API, CLI & MCP)**:
  - REST API: `POST /api/pdf/pages/preview` và `POST /api/pdf/tools` (hỗ trợ `reorganize`, `redact`, `protect`, `giap_lai`, `normalize_a4`, `extract_tables`).
  - CLI: Mở rộng lệnh `python ai_tools_cli.py pdf-tools` với đầy đủ các đối số mới.
  - MCP Server: Mở rộng công cụ `pdf_quick_tools` trong `server/mcp_server.py`.
- **Bộ Kiểm Thử Chất Lượng (Quality Gate)**:
  - Bổ sung `tests/test_pdf_pro_studio.py` (4 tests) và `tests/test_pdf_advanced_toolkit.py` (6 tests).
  - Toàn bộ test suite đạt **215 bài kiểm thử vượt qua 100% (215 passed, 0 failed, 3 warnings)**.

### Phiên Bản 4.0.0 (Phase 4 — Enterprise Unified QA Diagnostic, System Integrity & Final Integration)
- **Động Cơ Chẩn Đoán & Sửa Chữa Hợp Nhất (Enterprise Unified QA Diagnostic & Self-Repair)**:
  - `tools/unified_qa_diagnostic.py`:
    - **PowerPoint Diagnostic (`PptxDiagnosticEngine`)**:
      - Chẩn đoán chuẩn Widescreen 16:9 (**`PPTX_INV_01`**): Phát hiện tỷ lệ slide 4:3 gây co cụt hoặc vỡ lề khi trình chiếu.
      - Chẩn đoán tràn viền và thiếu word wrap (**`PPTX_INV_04`**): Quét toàn bộ text frame có nội dung, phát hiện nguy cơ tràn mép slide khi thiếu cờ `word_wrap=True`.
      - Chẩn đoán khóa thanh Ribbon / Color Picker (**`PPTX_INV_09`**): Phát hiện định dạng in đậm hoặc màu nền gán trực tiếp lên `<a:defRPr>` gây khóa cứng thanh chọn màu PowerPoint.
      - Chẩn đoán tính khả chỉnh của biểu đồ (**`PPTX_INV_14`**): Xác minh gói bảng tính Excel nhúng kèm theo các biểu đồ OpenXML Native.
      - Cơ chế Safe Non-destructive Repair cho PPTX: Tự động kích hoạt `word_wrap=True` và dọn dẹp các thuộc tính định dạng phi chuẩn trên `<a:defRPr>` mà không làm xô lệch vị trí hay font chữ gốc.
    - **Sơ Đồ Kỹ Thuật & Tài Sản Đồ Họa Vector (`DiagramAndConversionEngine`)**:
      - Kiểm định phân cấp Draw.io XML thuần (**`MX_INV_01`**): Ngăn chặn việc bọc thẻ `<UserObject>` chứa raw code PlantUML/Mermaid vi phạm kiến trúc native mxCell.
      - Kiểm định cú pháp XML hợp thức (**`MX_INV_02`**): Phát hiện tệp sơ đồ bị cắt cụt thẻ hoặc lỗi cú pháp XML.
      - Kiểm định mặt nạ nhãn đường nối line (**`MX_INV_08`**): Bắt buộc các đường line có nhãn văn bản phải có lớp nền `labelBackgroundColor=#ffffff` chống đè nét dây.
    - **Ảnh Scan & Phục Hồi Dấu Mộc Đỏ (`ERR_CONV_005`)**:
      - Tự động đo đạc tỷ lệ điểm ảnh mộc đỏ trong không gian màu HSV, khuyến nghị kích hoạt bộ lọc red-stamp trước khi đưa vào OCR.
    - **Tự Động Sửa Chữa Văn Bản Word (DOCX Safe Repair)**:
      - Tự động tiêm thẻ `<w:p/>` vào các ô bảng vi phạm *The Last Paragraph Rule* (**`ERR_DOCX_001`**).
      - Tự động tiêm cờ `<w:cantSplit/>` và `<w:tblHeader/>` cho các bảng biểu nhiều trang (**`ERR_DOCX_003`**).
      - Kiểm định tính hợp thức của Track Changes XML (**`ERR_DOCX_008`**).
- **Hợp Nhất Giao Diện Dòng Lệnh Master CLI (`qa-audit`)**:
  - `cli/ai_tools.py` & root facade `ai_tools_cli.py`:
    - Đăng ký subcommand `qa-audit`:
      `python ai_tools_cli.py qa-audit [--docx D] [--xlsx X] [--pptx P] [--diagram G] [--all DIR] [--mode {diagnose,audit-fix}] [--in-place] [--format {terminal,json,sarif}] [-o OUT]`
    - Mã thoát CI/CD chuẩn mực: `0` (Clean 100%), `1` (Cảnh báo thẩm mỹ), `2` (Lỗi cấu trúc nghiêm trọng).
- **Bộ Kiểm Thử Chất Lượng Toàn Diện (Quality Gate)**:
  - Bổ sung `tests/test_unified_qa_diagnostic.py` (5 tests chuyên sâu).
  - Toàn bộ test suite đạt **193 tests passed 100% (0 failed, 3 warnings in 53s)**.
- **Nâng Cấp Kiến Trúc `ARCHITECTURE.md` Lên v4.0.0**:
  - Chuẩn hóa mô hình 5 tầng: Presentation, Interface/Protocol, Core Engine, Storage/Cache, External Tooling.
  - Hoàn thiện tài liệu ma trận bất biến và tiêu chuẩn Clean-Room MIT Zero-COM.

### Phiên Bản 3.8.0 (Phase 3 — Enterprise Grade: Legal Redline Diff, Vietnamese OCR Pipeline & Linux/Docker Zero-COM)
- **So Sánh Văn Bản Pháp Lý & AI Track Changes (Redline Diff Engine)**:
  - `core/docx/diff_engine.py` & root facade `docx_diff_engine.py`:
    - Thuật toán so khớp văn bản chi tiết cấp đoạn văn và từ ngữ (`SequenceMatcher` Token-level Diff) bảo toàn khoảng trắng và dấu câu.
    - Sinh mã chuẩn OpenXML `<w:ins>` và `<w:del>` với thẻ `<w:delText>` (thay vì `<w:t>`), tuân thủ nghiêm ngặt quy chuẩn **ERR_DOCX_008 (Valid Track Changes XML)**.
    - Quản lý định danh `w:id` là số nguyên dương tăng dần duy nhất, tác giả `w:author` và timestamp ISO-8601 UTC `w:date`.
    - Tích hợp chèn nhận xét AI Reviewer vào lề phải tài liệu: đóng gói trực tiếp `word/comments.xml`, cập nhật `[Content_Types].xml` và liên kết quan hệ trong `word/_rels/document.xml.rels`.
    - Bảo toàn cấu trúc bảng biểu, tự động áp dụng `<w:cantSplit/>`, `<w:tblHeader/>` và quy tắc *The Last Paragraph Rule* (`ERR_DOCX_001`).
- **Pipeline OCR Tiếng Việt & Bóc Tách Hóa Đơn/Chứng Từ Scan (Scan Recovery)**:
  - `core/converter/ocr_pipeline.py` & root facade `ocr_pipeline.py`:
    - Thuật toán tự động phát hiện văn bản scan (`is_scanned_pdf`): mật độ text < 20 ký tự/trang kết hợp diện tích ảnh raster > 70% diện tích trang.
    - Bộ lọc mộc đỏ HSV (**ERR_CONV_005 (Scan & Red-Stamp Recovery)**): quét dải sắc độ đỏ trong không gian màu HSV (`H <= 14` hoặc `H >= 240`, `S >= 50`, `V >= 50`) để trung hòa dấu mộc đỏ, chống đứt gãy đường viền bảng tính và loại bỏ nhiễu đen khi qua OCR.
    - Bộ ước lượng góc nghiêng (`detect_skew_angle`) và nắn thẳng ảnh scan tự động (`deskew_image`) dựa trên phương sai của hồ sơ chiếu ngang.
    - Điều phối Tesseract OCR với gói ngôn ngữ tiếng Việt (`vie+eng`), tự động fallback sang `eng` hoặc trả về cấu trúc lỗi chuẩn khi thiếu binary ngoại vi.
    - Hỗ trợ xuất kết quả bóc tách ra tệp Markdown (`.md`) hoặc tài liệu Word (`.docx`).
- **Động Cơ Đa Nền Tảng Linux / Docker (Zero-COM Dependency)**:
  - `core/converter/pdf_renderer.py` & root facade `pdf_renderer.py`:
    - Lớp trừu tượng `UniversalPdfRenderer` với chuỗi fallback ưu tiên: Win32 COM (trên Windows có MS Office, kiểm tra an toàn không khóa tiến trình qua `winreg`) ➔ LibreOffice Headless (`soffice --headless --convert-to pdf`) ➔ Fallback PyMuPDF / Python.
    - Cung cấp các hàm tiện ích: `convert_docx_to_pdf()`, `convert_pptx_to_pdf()`, `convert_xlsx_to_pdf()`, `convert_to_pdf()`.
  - `Dockerfile` & `docker-compose.yml`:
    - Đóng gói container production Debian Bookworm với đầy đủ `libreoffice`, `tesseract-ocr`, `tesseract-ocr-vie`, `chromium`, OpenJDK JRE cho PlantUML, font chữ Liberation / DejaVu và Web Studio Server port 8000.
- **Mở Rộng Giao Diện & Bộ Công Cụ Chẩn Đoán (CLI, MCP, Web REST, Diagnostic)**:
  - CLI: Bổ sung lệnh `docx-diff` (hỗ trợ cả DOCX Redline Diff và JSON snapshot diff) và `ocr-scan` (bóc tách ảnh/PDF scan) trong `ai_tools_cli.py` và `cli/ai_tools.py`.
  - MCP Server: Đăng ký 2 công cụ mới `doc_redline_diff` và `doc_ocr_scan` trong `server/mcp_server.py` (tổng cộng 11 tools chuẩn MCP JSON-RPC 2.0).
  - REST API: Bổ sung 2 endpoints `POST /api/docx/diff` và `POST /api/ocr/scan` trong `server/app.py`.
  - Unified QA Diagnostic: Cập nhật kiểm tra quy chuẩn `ERR_DOCX_008` (xác thực tính hợp lệ của thẻ `<w:ins>`, `<w:del>`, `w:id`, `w:date`, `w:delText`) trong `tools/unified_qa_diagnostic.py`.
- **Bộ Kiểm Thử Chất Lượng (Quality Gate)**:
  - Mở rộng thêm 25 bài kiểm thử mới: `tests/test_docx_diff_redline.py` (10 tests), `tests/test_ocr_pipeline.py` (8 tests), `tests/test_pdf_renderer.py` (6 tests), và mở rộng `tests/test_mcp_server.py`.
  - Toàn bộ test suite đạt **188 bài kiểm thử vượt qua 100% (188 passed, 0 failed, 3 warnings in 51s)**.

### Phiên Bản 3.6.0 (Phase 2 — Batch Mail Merge Studio, Native OpenXML Charts & Enterprise Business Layouts)
- **Phân Hệ Trộn Thư Hàng Loạt (Batch Mail Merge Studio)**:
  - `core/docx/mail_merge.py` & root facade `docx_mail_merge.py`:
    - Quét và trích xuất token `{{Token}}` từ văn bản Word mẫu (`inspect_template_tokens`).
    - Nạp nguồn dữ liệu linh hoạt: Excel (`.xlsx`), file JSON (`.json`) hoặc danh sách đối tượng trực tiếp.
    - Bảo toàn số 0 ở đầu (`ERR_XLSX_007`), định dạng ngày tháng `YYYY-MM-DD` / `DD/MM/YYYY`, số tiền tệ.
    - Hỗ trợ khối logic điều kiện: `{{IF:Field=Val}}...{{ENDIF}}`, `{{IF:Field!=Val}}...{{ENDIF}}`, `{{IF:Field>Val}}...{{ENDIF}}`.
    - Hỗ trợ đặt tên file động: `HDLD_{{MaNV}}_{{HoTen}}.docx` và đóng gói nén toàn bộ tệp kết quả thành kho lưu trữ `.zip`.
- **Động Cơ Biểu Đồ Native OpenXML (`<c:chart>`) Cho PowerPoint & Word**:
  - `core/pptx/charts.py` & root facade `pptx_charts.py`:
    - Sinh thẻ chuẩn OpenXML `<c:chart>` cho 5 loại biểu đồ: Column (Clustered), Bar (Horizontal), Line (Markers), Pie (2D), Doughnut.
    - Nhúng workbook Excel ẩn (`embedded_xlsx`) cho phép chuột phải chọn "Edit Data" trực tiếp trong PowerPoint.
    - Tuân thủ nghiêm ngặt **PPTX_INV_14**: thiết lập màu sắc trục toạ độ và số liệu thuần đen `#000000` ở cấp run-level (`<a:rPr>`).
    - Các hàm cốt lõi: `add_chart_to_slide()`, `add_chart_slide()`, `build_chart_data()`.
  - `core/docx/charts.py` & root facade `docx_charts.py`:
    - Sinh mã vector SVG thuần túy cho 5 loại biểu đồ.
    - Kết xuất ảnh phân giải cao PNG qua headless browser (Edge/Chromium).
    - Khóa tỷ lệ và lề in ấn chuẩn A4 $\le 14\text{cm}$ (**ERR_DOCX_005**) kèm caption căn giữa in nghiêng.
  - `core/converter/markdown.py` & `markdown_converter.py`:
    - Bổ sung hỗ trợ fenced code block ````chart` chứa JSON/YAML spec, tự động render và nhúng biểu đồ vào tài liệu Word DOCX.
- **Mở Rộng 4 Bố Cục Slide Trình Diễn Doanh Nghiệp (Enterprise Business Layouts)**:
  - `core/pptx/writer.py` & root facade `pptx_writer.py`:
    - `add_kpi_summary_slide`: 4 thẻ chỉ số lớn kèm huy hiệu xu hướng `▲` (xanh ngọc) / `▼` (đỏ thẫm) và mô tả ngữ cảnh.
    - `add_timeline_roadmap_slide`: Trục thời gian ngang liên tục, các node đánh dấu tròn (01, 02...) và thẻ cột mốc chi tiết.
    - `add_swot_matrix_slide`: Lưới 2x2 ma trận SWOT cân đối (Strengths xanh dương, Weaknesses đỏ, Opportunities xanh lá, Threats hổ phách).
    - `add_process_chevron_slide`: Tiến trình tuần tự 3 đến 6 bước với mũi tên kết nối liên tục (`→`).
    - Bộ điều phối `write_pptx_from_spec`: tích hợp liền mạch các kiểu slide `chart`, `kpi_summary`, `timeline`, `swot`, `process`.
  - `core/pptx/docx_to_pptx.py` & root facade `docx_to_pptx.py`:
    - Bổ sung bộ phân loại ngữ nghĩa `_detect_business_layout()` tự động nhận diện tiêu đề và nội dung để sinh slide SWOT, Timeline, KPI Summary, hoặc Process.
- **Mở Rộng Giao Diện & Kết Nối (CLI, MCP, Web REST)**:
  - CLI: Bổ sung lệnh `docx-mail-merge` trong `ai_tools_cli.py` và `cli/ai_tools.py` (hỗ trợ cờ `--inspect`, `--pattern`, `--zip`).
  - Model Context Protocol: Đăng ký công cụ `doc_mail_merge` trong `server/mcp_server.py`.
  - Web REST API: Thêm endpoints `POST /api/mail-merge` và `POST /api/mail-merge/inspect` trong `server/app.py`.
- **Bộ Kiểm Thử Chất Lượng (Quality Gate Suite Expansion)**:
  - Tạo mới `tests/test_mail_merge.py` (4 tests), `tests/test_native_charts.py` (8 tests), `tests/test_pptx_business_layouts.py` (6 tests).
  - Nâng tổng số bài kiểm thử tự động toàn dự án lên **163 passed tests (100% Pass Rate)**.

### Phiên Bản 3.5.0 (Phase 1 — Administrative Presets, Vietnamese Typography Sanitizer, Print-Ready Excel & Everyday PDF Toolkit)
- **Bộ Preset Văn Bản Hành Chính (Nghị định 30/2020/NĐ-CP) & Luận Văn Học Thuật**:
  - `presets/nd30_hanh_chinh.style.yaml`: Quy định kích thước lề A4 chuẩn (Trên: 20mm, Dưới: 20mm, Trái: 30mm, Phải: 20mm), phông Times New Roman 13–14pt, giãn dòng 1.3 lines.
  - `presets/academic_thesis.style.yaml`: Lề đóng gáy trái 35mm, giãn dòng 1.5 lines, bảng màu Thesis Blue.
  - `core/docx/presets.py` & root facade `docx_presets.py`: Cung cấp `apply_preset_to_document()` và `build_administrative_header()` tạo bảng 2 cột ẩn viền cho Cơ quan chủ quản và Quốc hiệu / Tiêu ngữ tuân thủ nghiêm ngặt OpenXML invariants (`cantSplit`, `tblHeader`, `<w:p>` termination).
- **Bộ Sanitize Tiếng Việt & Dọn Dẹp Ký Tự Lạ Nâng Cao (`VietnameseTypographySanitizer`)**:
  - Tích hợp trong `core/docx/advanced_engine.py` (và root facade `docx_advanced_engine.py`).
  - Chuẩn hóa Unicode NFC (loại bỏ hoàn toàn lỗi vỡ dấu thanh tổ hợp tiếng Việt).
  - Chuẩn hóa khoảng trắng dấu câu (không khoảng trắng trước `, . ; : ! ?`, đúng 1 khoảng trắng sau dấu câu).
  - Ôm sát ngoặc đơn `( text )` ➔ `(text)` và ngoặc kép `“ quote ”` ➔ `“quote”`.
  - Chuẩn hóa En-dash (`–`) cho dải niên độ/số liệu: `2020-2026` ➔ `2020 – 2026`.
  - `sanitize_vietnamese_typography(doc)` duyệt đệ quy toàn bộ đoạn văn, bảng biểu, headers/footers.
- **Chuẩn Hóa In Ấn Bảng Tính Excel (`ERR_XLSX_008: Print-Ready Page Setup`)**:
  - `core/xlsx/writer.py` & `xlsx_writer.py`: Hàm `apply_worksheet_print_setup()` và `UniversalExcelEngine.apply_print_setup()`.
  - Thiết lập `fitToPage = True`, `fitToWidth = 1`, `fitToHeight = 0` (ngăn chặn triệt để hiện tượng vỡ bảng và tràn cột sang trang ngang).
  - Tự động xoay Landscape nếu số cột > 7, Portrait nếu $\le 7$.
  - Tự động phát hiện dòng Header qua Semantic Anchors và kích hoạt `ws.print_title_rows` lặp lại tiêu đề trên mọi trang in.
  - Tích hợp kiểm định và sửa chữa không phá hủy trong `tools/unified_qa_diagnostic.py`.
- **Bộ Công Cụ PDF Hàng Ngày (Everyday PDF Toolkit)**:
  - `core/converter/pdf_toolkit.py` & root facade `pdf_toolkit.py`:
    - `merge_pdfs`: Gộp tuần tự danh sách tệp PDF.
    - `split_pdf`: Tách PDF thành từng trang đơn lẻ hoặc theo dải trang (`1-3, 4-5`).
    - `rotate_and_clean_pdf`: Xoay trang theo góc (90°, 180°, 270°) và nhận diện/loại bỏ trang trắng.
    - `stamp_watermark_pdf`: Đóng dấu chìm watermark bán trong suốt, tự động dò tìm và nhúng TrueType font hệ thống hỗ trợ tiếng Việt Unicode 100%.
    - `compress_pdf`: Tối ưu hóa xref, deflation và dọn dẹp rác bộ nhớ giảm dung lượng PDF.
- **Mở Rộng Giao Tiếp & Giao Diện (CLI, MCP, Web REST)**:
  - CLI: Bổ sung các lệnh `docx-apply-preset`, `docx-sanitize-vietnamese`, `pdf-tools` trong `ai_tools_cli.py` và `cli/ai_tools.py`.
  - Model Context Protocol: Đăng ký công cụ `docx_apply_preset` và `pdf_quick_tools` trong `server/mcp_server.py`.
  - Web API REST: Thêm endpoints `GET /api/presets`, `POST /api/docx/preset`, `POST /api/pdf/tools`, `POST /api/xlsx/print-setup` trong `server/app.py`.
- **Bộ Kiểm Thử Chất Lượng Nâng Cấp**:
  - Tạo mới `tests/test_nd30_preset.py` (5 tests), `tests/test_xlsx_print_setup.py` (5 tests), `tests/test_pdf_toolkit.py` (5 tests).
  - Mở rộng tổng số bài kiểm thử tự động lên **145 passed tests (100% Pass Rate)**.

### Phiên Bản 3.4.0 (Phiên Bản Nền Tảng — OpenXML List & Numbering Engine + Unified Diagram Bridge + 108 Test Suites)
- **Động cơ Quản lý Danh sách Độc lập (`docx_lists.py`)**:
  - Đăng ký cấu trúc OpenXML `abstractNum` chuẩn ECMA-376 (ID 2001 cho số thứ tự 3 cấp `1.` -> `a.` -> `i.`; ID 2002 cho bullet 3 cấp `•` -> `○` -> `■`).
  - Cấp phát `numId` động độc lập kèm `<w:startOverride w:val="1"/>`, tự động reset danh sách về 1 khi bắt đầu mục mới.
  - **Cơ chế Smart Continuation**: Duy trì liên tục số thứ tự cấp cha ngay cả khi có ghi chú hoặc sub-items xen kẽ.
  - **Dynamic Native Word Numbering**: Khi mở tài liệu trong Word, nhấn Enter sẽ tự động nhảy số tiếp theo thay vì văn bản tĩnh.
  - **Task Lists Checkbox**: Chuyển đổi mượt mà `- [ ]` sang `☐` (ballot box) và `- [x]` / `* [X]` sang `☑` (checked box) in đậm.
  - **Hanging Indentation Invariants**: Áp dụng chuẩn thụt lề dxa `<w:ind w:left="..." w:hanging="360"/>` (cấp 0: 720, cấp 1: 1440, cấp 2: 2160).
  - **Strict Guard**: Tuyệt đối không tự ý sinh số thứ tự hoặc bullet cho các đoạn văn thông thường.
- **Cơ chế Sinh Sơ Đồ Đa Định Dạng Thống Nhất (`core/diagrams/bridge.py` & `diagram_bridge.py`)**:
  - Hợp nhất 4 động cơ đồ họa (Mermaid, PlantUML, Draw.io XML, Canvas Spec) qua giao diện chung.
  - Biên dịch trực tiếp khối mã sơ đồ fenced trong Markdown (` ```mermaid `, ` ```plantuml `, ` ```drawio `, ` ```canvas `, ` ```spec `) thành hình ảnh PNG chất lượng cao nhúng vào DOCX, PPTX và PDF.
  - Hỗ trợ cú pháp liên kết `![Caption](spec.json)` tự động dò tìm và render spec đồ họa thành ảnh tương ứng.
  - Khóa tỷ lệ khung hình chuẩn A4 (`ERR_DOCX_005`: $W \le 14\text{cm}, H \le 18\text{cm}$) và gắn chú thích in nghiêng $9.5\text{pt}$ căn giữa.
  - Bộ nhớ đệm Content-Hash Caching (MD5) chống render trùng lặp, tăng tốc tối đa khi xử lý tài liệu lớn.
- **Tích hợp sâu vào `markdown_converter.py` & `docx_writer.py`**:
  - `markdown_converter.py`: Tích hợp `ListContextTracker` nhận diện thụt đầu dòng; tự động reset danh sách khi gặp tiêu đề `#`, code blocks, ảnh, bảng biểu, ngắt trang, hoặc văn bản thường.
  - **Hỗ trợ Danh sách trong Bảng (Table Cell Lists)**: Phân giải các dòng ngắt `<br>` hoặc `\n` trong ô bảng biểu thành các đoạn danh sách chuẩn OpenXML, giữ vững quy tắc *The Last Paragraph Rule* (`ERR_DOCX_001`).
  - `docx_writer.py`: Hỗ trợ trường `list_type`, `list_level`, `list_restart`, `is_checked` trong paragraph AST và cell paragraphs.
- **Mở Rộng Test Suites Lên 108 Bài Kiểm Thử (100% Pass Rate)**:
  - Tạo mới `tests/test_docx_lists.py` (22 bài test chuyên sâu cho OpenXML list engine).
  - Tạo mới `tests/test_docx_extended.py` (25 bài test mở rộng kiểm tra các bất biến bố cục, màu sắc, bảng biểu, lề trang, Math, và template patching).
  - Tổng số bài test đạt **108 bài kiểm thử tự động**, đạt tỷ lệ đạt 108/108 (100% Pass Rate).

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

- **Tổng số ca kiểm thử**: 220 bài test tự động.
- **Trạng thái thực thi**: 220/220 bài test tự động (**100% Pass Rate**).
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
