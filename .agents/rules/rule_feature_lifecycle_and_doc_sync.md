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

---

## 3. Kiểm Chuẩn Bắt Buộc (Quality Gate Audit)

Trước khi bàn giao kết quả cho người dùng, Agent **BẮT BUỘC** thực hiện kiểm tra 3 lớp:
1. **Lớp 1 - Kiểm thử tự động**: Chạy toàn bộ test suite để đảm bảo không phát sinh regression:
   ```bash
   python -m unittest discover -s tests
   ```
   *Bắt buộc 100% Pass Rate.*
2. **Lớp 2 - Kiểm tra đồng bộ phiên bản**: Nếu có đổi version, chạy kiểm tra đối soát 9 điểm neo:
   ```bash
   python tools/bump_version.py --check
   ```
   *Bắt buộc trả về `[OK] Invariant ERR_VER_001 PASSED`.*
3. **Lớp 3 - Khớp số liệu tài liệu**: Số lượng test và số phiên bản được ghi trong `CHANGELOG.md`, `ARCHITECTURE.md` và `README.md` phải khớp 100% với con số thực tế chạy được ở Lớp 1.

---

## 4. Các Hành Vi Bị Nghiêm Cấm (Anti-Patterns)

1. ❌ **"Code xong báo xong" mà quên cập nhật tài liệu**: Báo hoàn tất cho người dùng mà không cập nhật `CHANGELOG.md` và `ARCHITECTURE.md`.
2. ❌ **Đoán mò số lượng bài kiểm thử**: Tự bịa hoặc đoán số lượng test ghi vào tài liệu thay vì chạy lệnh `python -m unittest discover -s tests` để lấy số thực tế.
3. ❌ **Nâng version thủ công rải rác**: Tự sửa tay số phiên bản ở 1-2 file dẫn đến Version Desync giữa Frontend, Backend và CLI.
4. ❌ **Bỏ quên Frontend Build**: Sửa đổi tính năng Web Studio mà không chạy `npm run build` để cập nhật `dist/`.
