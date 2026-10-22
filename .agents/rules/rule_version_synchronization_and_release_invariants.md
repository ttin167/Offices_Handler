# Rule: Version Synchronization & Dual-Surface Release Invariants

> **Scope**: Applies every time a version milestone, new feature, bug fix, or release is made in `antigravity-doc-handler`.
> **Invariant Code**: `ERR_VER_001` (**Version Desync Invariant**)
> **Trigger**: Version bump, release mới, cập nhật phiên bản, package.json, changelog, terminal/FE desync.

---

## 1. Nguyên Tắc Cốt Lõi: Đồng Bộ Tuyệt Đối Hai Bề Mặt (Dual-Surface Synchronization)

Trong một hệ thống Full-stack tích hợp Terminal CLI, Dev Launcher, Backend Python, MCP Server, và Vite Web Studio, **tuyệt đối không được phép xảy ra hiện tượng phân mảnh phiên bản (Version Desync)** giữa các tầng:
- Terminal/CLI hiển thị một phiên bản cũ (ví dụ: `v4.0.0` trong `node dev_launcher.js` hoặc `package.json`).
- Tài liệu ghi nhận một phiên bản khác (ví dụ: `v4.2.0` trong `CHANGELOG.md`).
- Frontend hiển thị một phiên bản thứ ba (ví dụ: `v4.0` trong `index.html`).

Mọi phiên bản phát hành **BẮT BUỘC** phải được đồng bộ hóa đồng thời trên toàn bộ 9 điểm neo (Anchors) của hệ thống.

---

## 2. Danh Mục 9 Điểm Neo Bắt Buộc Đồng Bộ (The 9 Synchronization Anchors)

Mỗi khi nâng cấp phiên bản `X.Y.Z`, toàn bộ các tệp sau bắt buộc phải mang đúng chuỗi phiên bản:

| # | Tệp Mục Tiêu | Vị Trí / Thuộc Tính | Mục Đích |
|---|---|---|---|
| 1 | `package.json` | `"version": "X.Y.Z"` | Định danh gói npm, hiển thị khi `npm run build` |
| 2 | `package-lock.json` | `version` root & `packages[""]` | Khóa phiên bản dependency cây npm |
| 3 | `pyproject.toml` | `version = "X.Y.Z"` | Khóa phiên bản gói phân phối Python |
| 4 | `dev_launcher.js` | `let pkgVersion = 'X.Y.Z';` | Banner khởi động Terminal `node dev_launcher.js` |
| 5 | `server/app.py` | `APP_VERSION = "X.Y.Z"` | Phiên bản phục vụ REST `/api/health` & `/api/info` |
| 6 | `server/mcp_server.py` | `SERVER_VERSION = "X.Y.Z"` | Định danh phiên bản giao thức Model Context Protocol |
| 7 | `core/env_doctor.py` | `DOCTOR_VERSION = "X.Y.Z"` | Banner chẩn đoán `ai_tools_cli.py doctor` |
| 8 | `index.html` | `.brand-subtitle` (`Office Studio vX.Y.Z`) | Hiển thị tĩnh trên Header khi tải trang ban đầu |
| 9 | `README.md` | Badges (`Release-vX.Y.Z_Production`, Tests, MCP Tools), Highlights header, Link `CHANGELOG.md (v1.0 -> vX.Y.Z)` | Huy hiệu, số liệu kiểm thử & điểm nhấn phát hành trên kho mã nguồn |

Đồng thời cập nhật tương ứng tại:
- `CHANGELOG.md`: Mục phiên bản mới ở đầu phần 2 và bảng so sánh phần 1.
- `ARCHITECTURE.md`: Trường `Phiên bản: X.Y.Z` ở đầu tài liệu.

---

## 3. Quy Trình Thực Thi Bắt Buộc Khi Nâng Cấp Phiên Bản

### Bước 1: Kích Hoạt Đồng Bộ 1 Lệnh Duy Nhất (Single Command Bump)
Chạy lệnh đồng bộ tự động qua script đã chuẩn hóa:
```bash
python tools/bump_version.py <X.Y.Z>
# Hoặc qua lệnh npm:
npm run bump -- <X.Y.Z>
```
Script này tự động:
1. Ghi đè phiên bản mới đồng loạt lên 9 điểm neo trên.
2. Tự động kích hoạt `npm run build` để biên dịch lại bundle trong `dist/` phục vụ cho Web Studio production.

### Bước 2: Kiểm Tra Đối Soát Chất Lượng (Quality Gate Audit)
Bắt buộc chạy lệnh kiểm tra tính nhất quán (Audit Check):
```bash
python tools/bump_version.py --check
```
- Nếu trả về `[OK] Invariant ERR_VER_001 PASSED`: Đủ điều kiện xuất xưởng.
- Nếu trả về `[FAIL] Invariant ERR_VER_001 VIOLATION`: Dừng lại ngay lập tức, khắc phục triệt để điểm lệch và kiểm tra lại.

### Bước 3: Cơ Chế Client-Side Dynamic Synchronization
Trên giao diện Frontend:
- `src/main.js` tại hàm `checkServerHealth()` **BẮT BUỘC** đọc thuộc tính `version` từ response của endpoint `/api/health` để cập nhật động vào phần tử DOM `.brand-subtitle`:
  ```javascript
  if (data.version) {
    const brandSubtitle = document.querySelector('.brand-subtitle');
    if (brandSubtitle) {
      brandSubtitle.textContent = `Office Studio v${data.version}`;
    }
  }
  ```
Điều này đảm bảo ngay cả khi người dùng truy cập qua trình duyệt chưa xóa cache HTML, giao diện vẫn luôn hiển thị chính xác phiên bản mà backend đang chạy.

---

## 4. Các Hành Vi Bị Nghiêm Cấm (Anti-Patterns)

1. ❌ **Cấm nâng phiên bản cục bộ chỉ trong tài liệu**: Nghiêm cấm ghi nhận phiên bản mới vào `CHANGELOG.md` hay `ARCHITECTURE.md` mà bỏ quên các tệp runtime (`package.json`, `server/app.py`, `dev_launcher.js`).
2. ❌ **Cấm sửa giao diện FE mà không rebuild `dist/`**: Nghiêm cấm thay đổi phiên bản trên `index.html` hoặc `src/main.js` mà không chạy `npm run build`, khiến file bundle production bị lỗi thời.
3. ❌ **Cấm gán chuỗi version hardcode phân mảnh**: Khi tạo endpoint hoặc công cụ mới cần trả về version, luôn tham chiếu từ `APP_VERSION` hoặc đọc từ `package.json`, không hardcode chuỗi `4.0.0` rải rác trong code.
