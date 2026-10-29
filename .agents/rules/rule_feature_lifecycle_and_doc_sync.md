# Rule: Feature Lifecycle, Documentation Synchronization & Release Invariants

> **Scope**: Applies every time a feature, sub-system, API, CLI command, or bug fix is added or modified in `antigravity-doc-handler`.
> **Invariant Code**: `ERR_DOC_SYNC_001` (**Feature Completion & Documentation Invariant**)
> **Trigger Keywords**: Feature completion, hoàn tất tính năng, cập nhật changelog, architecture, doc sync, nâng cấp phiên bản.

---

## 1. Nguyên Tắc Cốt Lõi: Hoàn Tất Tính Năng Đi Liền Với Đồng Bộ Tài Liệu (Feature-Doc Parity)

Trong `antigravity-doc-handler`, một tính năng **CHƯA ĐƯỢC COI LÀ HOÀN TẤT** nếu chỉ có mã nguồn hoạt động và vượt qua test.
Bất kỳ khi nào hoàn tất cập nhật hoặc chỉnh sửa một tính năng (dù lớn hay nhỏ), AI Agent **BẮT BUỘC** phải thực hiện quy trình đồng bộ hóa tài liệu kỹ thuật trước khi thông báo hoàn tất cho người dùng:
1. Ghi nhận đầy đủ vào `CHANGELOG.md` theo chuẩn phân loại.
2. Cập nhật cấu trúc module, bảng điều phối (routing table) và số liệu kiểm thử trong `ARCHITECTURE.md`.
3. Nếu có nâng cấp phiên bản (Version Bump), bắt buộc chạy quy trình đồng bộ 9 điểm neo qua `python tools/bump_version.py <X.Y.Z>`.
4. Vượt qua Kiểm chuẩn 3 lớp (Quality Gate Audit).

---

## 2. Quy Định Cụ Thể Cho Từng Tài Liệu

### A. Đồng Bộ `CHANGELOG.md`
Mỗi khi tính năng hoàn tất, cập nhật vào section phiên bản hiện tại (hoặc tạo section phiên bản mới nếu nâng version):
- **Phân loại chuẩn**: Sử dụng các mục con:
  - `### Added`: Các tính năng mới, CLI sub-commands mới, API endpoints mới, hoặc test suites mới.
  - `### Changed`: Các thay đổi cấu trúc, tái cấu trúc logic, mở rộng tham số hoặc tối ưu hiệu năng.
  - `### Fixed`: Các bản vá lỗi, khắc phục ngoại lệ, vá bẫy kỹ thuật (edge cases).
- **Ghi nhận số lượng Test Suite**: Nêu rõ các test suite mới được viết kèm tổng số ca kiểm thử tự động của toàn hệ thống (ví dụ: `189 Automated Test Suites (100% Pass Rate)`).

### B. Đồng Bộ `ARCHITECTURE.md`
`ARCHITECTURE.md` là nguồn chân lý kiến trúc (Single Source of Truth). Bắt buộc rà soát và cập nhật:
1. **Header & Metadata**: Cập nhật số phiên bản và tổng số bài test ở đầu tài liệu.
2. **Bảng Phân Bổ Module & Sơ Đồ Cấu Trúc**: Nếu có module mới (ví dụ: `core/converter/table_extractor.py`), phải thêm vào bảng thư mục mã nguồn và sơ đồ phân tầng.
3. **Bảng Điều Phối Lệnh (CLI / API / MCP Routing Table)**: Bổ sung các subcommand CLI, endpoint REST hoặc tool MCP mới tương ứng.
4. **Mục Kiểm Định Chất Lượng**: Cập nhật số lượng bài test thực tế đạt 100% Pass Rate.

### C. Quy Trình Nâng Cấp Phiên Bản (Version Bump)
Khi nhiệm vụ có yêu cầu nâng phiên bản (ví dụ: `4.2.0` $\to$ `4.3.0` hoặc `4.2.1`):
1. **Khởi chạy script tự động**:
   ```bash
   python tools/bump_version.py <X.Y.Z>
   # Hoặc: npm run bump -- <X.Y.Z>
   ```
   *Script này tự động ghi đè 9 điểm neo (`package.json`, `package-lock.json`, `pyproject.toml`, `dev_launcher.js`, `server/app.py`, `server/mcp_server.py`, `core/env_doctor.py`, `index.html`, `README.md`) và chạy `npm run build` tạo lại bundle `dist/`.*
2. **Tạo Section Mới Trong `CHANGELOG.md`**: Ghi rõ ngày phát hành và danh mục thay đổi.
3. **Cập nhật Phiên Bản trong `ARCHITECTURE.md`**: Đổi số phiên bản ở đầu trang.

### D. Đồng Bộ Bộ Đôi Tài Liệu Cốt Lõi (Core Technical & User Documentation Triad: MD -> DOCX & PDF)
Bên cạnh `CHANGELOG.md` và `ARCHITECTURE.md`, hệ sinh thái quy định **BỘ ĐÔI TÀI LIỆU CỐT LÕI (CORE DOCS)** bắt buộc phải duy trì tính toàn vẹn và đồng bộ theo thời gian thực:
1. **Tài liệu Kiến trúc**: `docs/BAO_CAO_TIEN_HOA_KIEN_TRUC_V1_TO_V4.md` (Báo cáo tiến hóa kiến trúc đa tầng, ma trận so sánh các phiên bản v1 $\to$ v4+, các phân hệ Core và quy chuẩn bất biến).
2. **Tài liệu Hướng dẫn sử dụng**: `docs/HUONG_DAN_SU_DUNG_CHI_TIET_TUNG_TINH_NANG.md` (Cẩm nang thao tác cho người dùng cuối qua Web Studio, Desktop GUI, AI CLI, và MCP Server).

**Quy chuẩn Bộ Ba Định Dạng (Triad Invariant)**:
- Mỗi khi có tính năng mới, thay đổi API, bổ sung lệnh CLI hoặc cập nhật giao diện, AI Agent **BẮT BUỘC** cập nhật nội dung vào 2 file Markdown trên.
- Đồng thời, **BẮT BUỘC** recompile tự động ra cả bản Word (`.docx`) và PDF (`.pdf`) để bộ tài liệu luôn sẵn sàng in ấn/xuất bản chuẩn quốc tế.

**Bất Biến Đồng Bộ Ngữ Nghĩa Thân Tài Liệu (Semantic Content Parity Checklist — 5 Tiêu Chuẩn Bắt Buộc)**:
> [!CAUTION]
> Nghiêm cấm tuyệt đối hành vi "chỉ sửa version ở dòng tiêu đề đầu file" mà bỏ quên nội dung bên trong thân tài liệu (Semantic Body Desync). Mọi đợt cập nhật hoặc Version Bump bắt buộc tuân thủ 5 tiêu chuẩn:
1. **Khớp Phiên Bản Bảng So Sánh (Section 2)**: Tiêu đề bảng đối chiếu (`## 2. BẢNG MA TRẬN SO SÁNH TOÀN DIỆN (v1.0 BASELINE VS vX.Y.Z PRODUCTION GRADE)`) và tiêu đề cột phiên bản hiện tại bắt buộc phải mang đúng nhãn `vX.Y.Z Production Grade`. Không được để sót bất kỳ chuỗi phiên bản cũ nào ở tiêu đề bảng so sánh.
2. **Bổ Sung Dòng Đối Chiếu Phân Hệ Mới**: Mọi phân hệ, động cơ hoặc tính năng mới ra mắt trong phiên bản bắt buộc phải có các dòng phân tích đối chiếu chuyên sâu (Baseline vs Hiện tại) trong Bảng ma trận so sánh Phần 2.
3. **Chương Phân Tích Tính Năng Chuyên Sâu**: Bắt buộc có một Chương/Mục riêng biệt phân tích chi tiết kiến trúc, API/CLI, luồng dữ liệu và mã bất biến kỹ thuật liên quan của phiên bản đó (ví dụ: `## 10. TIẾN HÓA 6 PHÂN HỆ DOANH NGHIỆP MỚI TRONG PHIÊN BẢN vX.Y.Z`).
4. **Khớp Số Liệu Kiểm Thử & MCP Catalog**: Số lượng bài test tự động (ví dụ: `314+ Automated Tests` / `314/314 Passed`) và số lượng công cụ MCP (ví dụ: `28 MCP Tools Catalog`) ghi trong thân tài liệu bắt buộc phải khớp 100% với kết quả chạy kiểm thử thực tế và số tool đăng ký trong `server/mcp_server.py`.
5. **Đồng Bộ Đánh Giá Hiệu Năng & Kết Luận**: Chương Đánh giá hiệu năng và Kết luận (Chương cuối) bắt buộc phải cập nhật số liệu đo lường, tỷ lệ test pass 100%, và tổng kết chính thức của phiên bản `vX.Y.Z`.

- Lệnh đồng bộ tự động 1-click (MD $\to$ DOCX & PDF):
  ```bash
  python tools/sync_core_docs.py --build
  ```
- Lệnh kiểm tra đối soát tính đồng bộ (Kiểm tra cả mtime lẫn Semantic Body Parity):
  ```bash
  python tools/sync_core_docs.py --check
  ```

### E. Đồng Bộ `README.md` & Hệ Thống Kỹ Năng (README & Skills Synchronization)
`README.md` là bộ mặt công khai của repository, chứa các huy hiệu (badges), bảng năng lực kỹ năng của AI Agent và điểm nhấn phiên bản:
1. **Khi nâng cấp kỹ năng hoặc thêm skill mới (`.agents/skills/`)**:
   - Cập nhật số lượng kỹ năng trên badge `Antigravity_Skills-N_Specialized`.
   - Bổ sung/cập nhật dòng tương ứng trong bảng `## 🧠 Antigravity AI Skills & Customizations Suite`.
   - Cập nhật danh mục kỹ năng trong [AGENTS.md](AGENTS.md).
2. **Khi nâng cấp phiên bản (Version Bump)**:
   - Cập nhật badge phiên bản `Release-vX.Y.Z_Production`, số lượng test cases tự động, số lượng công cụ MCP.
   - Cập nhật tiêu đề mục `### 🆕 Highlights in vX.Y.Z Production Grade`.
   - Cập nhật liên kết chuyển tiếp tới `CHANGELOG.md (v1.0 → vX.Y.Z)`.
3. **Lệnh đồng bộ 1-click**:
   ```bash
   python tools/bump_version.py --sync-readme
   ```

---

## 3. Kiểm Chuẩn Bắt Buộc (Quality Gate Audit)

Trước khi bàn giao kết quả cho người dùng, Agent **BẮT BUỘC** thực hiện kiểm tra 4 lớp:
1. **Lớp 1 - Kiểm thử tự động**: Chạy toàn bộ test suite để đảm bảo không phát sinh regression:
   ```bash
   python -m unittest discover -s tests
   ```
   *Bắt buộc 100% Pass Rate.*
2. **Lớp 2 - Kiểm tra đồng bộ phiên bản & README**: Chạy kiểm tra đối soát 9 điểm neo:
   ```bash
   python tools/bump_version.py --check
   ```
   *Bắt buộc trả về `[OK] Invariant ERR_VER_001 PASSED`.*
3. **Lớp 3 - Khớp số liệu tài liệu**: Số lượng test, số lượng kỹ năng, số lượng MCP tools và số phiên bản được ghi trong `CHANGELOG.md`, `ARCHITECTURE.md` và `README.md` phải khớp 100% với con số thực tế chạy được ở Lớp 1.
4. **Lớp 4 - Tính toàn vẹn Core Docs Triad & Ngữ Nghĩa Thân Tài Liệu**: Chạy `python tools/sync_core_docs.py --check` (kiểm tra cả mtime đồng bộ DOCX/PDF lẫn 5 tiêu chuẩn Semantic Body Parity). Trả về Exit Code 1 nếu phát hiện bất kỳ sự lệch pha nào trong thân tài liệu.

---

## 4. Các Hành Vi Bị Nghiêm Cấm (Anti-Patterns)

1. ❌ **"Code xong báo xong" mà quên cập nhật tài liệu**: Báo hoàn tất cho người dùng mà không cập nhật `CHANGELOG.md`, `ARCHITECTURE.md` và `README.md`.
2. ❌ **Quên cập nhật `README.md` khi nâng cấp kỹ năng hoặc nâng version**: Để `README.md` mang số liệu cũ, badge lệch hoặc thiếu vắng các kỹ năng mới được trang bị trong `.agents/skills/`.
3. ❌ **Đoán mò số lượng bài kiểm thử**: Tự bịa hoặc đoán số lượng test ghi vào tài liệu thay vì chạy lệnh kiểm thử thực tế để lấy số chính xác.
4. ❌ **Nâng version thủ công rải rác**: Tự sửa tay số phiên bản ở 1-2 file dẫn đến Version Desync giữa Frontend, Backend và CLI.
5. ❌ **Bỏ quên Frontend Build**: Sửa đổi tính năng Web Studio mà không chạy `npm run build` để cập nhật `dist/`.
6. ❌ **"Chỉ sửa số version ở tiêu đề" (Title-Only Version Bump)**: Thay đổi nhãn phiên bản ở dòng đầu tệp Markdown hoặc chỉ sửa mtime mà bỏ quên bảng ma trận so sánh, số lượng kiểm thử, số công cụ MCP và chương tính năng trong thân tài liệu (Semantic Body Desync).
