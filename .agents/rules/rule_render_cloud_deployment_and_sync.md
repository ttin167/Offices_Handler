# Rule: Render Cloud 24/7 Deployment & Single-Port Synchronization Invariants (RND_INV_01–08)

> **Scope**: Mandatory rule applied whenever modifying deployment configurations (`render.yaml`, `Dockerfile`), building frontend bundles, pushing releases to Render.com cloud, or managing cloud container resources (512MB RAM tier).
> Reference Guide: [RENDER_DEPLOYMENT_GUIDE.md](../../RENDER_DEPLOYMENT_GUIDE.md)

---

## 🎯 Core Deployment Mental Model

Render.com Cloud runs in an ephemeral container environment with a **512MB RAM ceiling** (Free Tier). To ensure 24/7 zero-downtime, zero-OOM crashes (Exit code 137), and fast deployments, every release must strictly comply with the **8 Render Invariants (`RND_INV_01`–`RND_INV_08`)**.

```
       Internet / Người Dùng (HTTPS)
                    │
                    ▼
           [ Render Cloud Proxy ]
            (Cấp phát cổng $PORT)
                    │
                    ▼
      [ Docker Container: Python Web Server ]
               (0.0.0.0:$PORT)
                    │
         ┌───────────┴───────────┐
         ▼                       ▼
   [ /api/* Routes ]       [ Static SPA Routes ]
   - /api/health           - / (dist/index.html)
   - /api/ba/*             - /assets/*.js, *.css
   - /api/convert          - SPA Fallback Route
   - /api/jobs/* (SSE)
```

---

## 🛡️ Render Invariants Matrix (RND_INV_01 – RND_INV_08)

| Mã Invariant | Tên Quy Chuẩn | Yêu Cầu Kỹ Thuật Bắt Buộc | Rủi Ro Nếu Vi Phạm |
|---|---|---|---|
| `RND_INV_01` | **Single-Port Routing & Dynamic `$PORT` Binding** | `server/app.py` và `web_api_server.py` bắt buộc đọc biến môi trường `PORT = int(os.environ.get("PORT", 10000))` và bind vào `HOST = "0.0.0.0"`. Static SPA (`/`, `/assets/*`) phải phục vụ từ thư mục `dist/` do Python HTTP Server trực tiếp quản lý. | Render báo `Port Binding Timeout` hoặc `Connection Refused` do proxy không tìm thấy cổng. |
| `RND_INV_02` | **512MB RAM Concurrency & Heap Guard** | Khóa cứng `MAX_JOB_WORKERS=1` trong `render.yaml` và `Dockerfile` để tuần tự hóa hàng đợi CPU/RAM nặng (LibreOffice, Chromium, Java). Cấu hình bắt buộc JVM: `JAVA_TOOL_OPTIONS="-Xmx192m -DPLANTUML_LIMIT_SIZE=16384 -Dfile.encoding=UTF-8"`. | Tràn RAM 512MB khiến hệ thống bị Linux Kernel OOM Killer tiêu diệt tức thì (`Exit code 137`). |
| `RND_INV_03` | **Multi-Stage Build & Frontend Pre-Compilation** | Dockerfile bắt buộc dùng mô hình 2 tầng (Stage 1: `node:18-alpine` chạy `npm run build` tạo `dist/`; Stage 2: `python:3.11-slim-bookworm` sao chép `dist/` vào `/app/dist`). Nghiêm cấm chạy Vite Dev Server trên môi trường production. | Kích thước Docker image phình to > 3GB, lãng phí tài nguyên và làm chậm tốc độ khởi động. |
| `RND_INV_04` | **Headless Sandbox Security & Binary Presence** | Chromium bắt buộc gọi kèm các cờ `--no-sandbox --disable-dev-shm-usage --headless`. Container phải cài đủ 5 công cụ lõi: `libreoffice`, `tesseract-ocr` (vie+eng), `openjdk-17-jre-headless`, `graphviz`, `chromium`, và hệ font Noto/DejaVu/Liberation. `plantuml.jar` phải có cơ chế curl fallback. | Sơ đồ PlantUML (Class, C4, Use Case) crash do thiếu `dot`; sơ đồ Mermaid/Draw.io lỗi do thiếu Chromium; tiếng Việt bị vỡ ô vuông `[]`. |
| `RND_INV_05` | **Ephemeral Disk & 2-Hour Auto-Purge** | Đĩa lưu trữ của Render Free là tạm thời (Ephemeral Disk). Bắt buộc áp dụng trần tải lên `MAX_UPLOAD_SIZE_MB=50` và tự động kích hoạt `cleanup_old_temp_files(max_age_seconds=7200)` sau mỗi tác vụ để dọn sạch tệp tạm quá 2 giờ. | Đĩa tạm phình to làm cạn kiệt Inode và bộ nhớ chia sẻ, khiến server từ chối ghi tệp mới. |
| `RND_INV_06` | **Docker Layer Cache Preservation Policy (~1GB Cache)** | **CHÍNH SÁCH CACHE**: Dung lượng cache ~800MB–1GB trên Render là các tầng Docker Layer được tái sử dụng (APT packages, Python wheels, Node base). **TUYỆT ĐỐI KHÔNG XÓA CACHE** trong các lần push cập nhật mã thông thường để giữ thời gian build chỉ mất 30–60 giây. **CHỈ XÓA CACHE** khi: thay đổi gói APT trong Dockerfile, nâng cấp phiên bản runtime nền, hoặc layer bị hỏng. | Xóa cache làm mất toàn bộ layers, buộc tải và cài lại từ đầu 15–20 phút, dễ dẫn đến timeout hủy build. |
| `RND_INV_07` | **Zero-COM Linux Headless Fallback** | Mọi phân hệ xử lý tài liệu, in ấn và đồ họa phải có chuỗi dự phòng Linux: không dùng `win32com`, không dùng `win32print`. Tự động fallback sang LibreOffice CLI, CUPS / `lpstat`, hoặc Virtual Spool Buffer (`JOB-...`). | Tính năng in ấn hoặc chuyển đổi tài liệu ném ngoại lệ `ModuleNotFoundError: win32print` trên Linux. |
| `RND_INV_08` | **Pre-Push 4-Gate Quality Verification** | Trước khi push git lên branch liên kết Render (`autoDeploy: true`), bắt buộc chạy 4 cổng kiểm định cục bộ: 1) `npm run build`; 2) `python tools/sync_core_docs.py --check`; 3) `python -m pytest tests/`; 4) Kiểm tra tính nhất quán giữa `render.yaml` và `Dockerfile`. | Đẩy mã lỗi lên repo khiến Render tự động deploy thất bại hoặc làm sập web đang hoạt động 24/7. |

---

## 📋 Checklist Khi Cập Nhật Mã Nguồn Lên Render

Khi chuẩn bị đưa bản cập nhật mới lên Render:

- [ ] **Bước 1**: Biên dịch frontend bằng `npm run build` — xác nhận `dist/index.html` và assets được tạo sạch sẽ không lỗi.
- [ ] **Bước 2**: Thẩm định tính đồng bộ tài liệu qua `python tools/sync_core_docs.py --check` — đạt kết quả `[OK]`.
- [ ] **Bước 3**: Kiểm tra toàn bộ test suite `python -m pytest tests/` — đạt 100% test passed.
- [ ] **Bước 4**: Kiểm tra cấu hình `render.yaml` — bảo đảm `HOST=0.0.0.0`, `MAX_JOB_WORKERS=1`, `JAVA_TOOL_OPTIONS` đầy đủ cờ hạn chế heap 192MB.
- [ ] **Bước 5**: Giữ nguyên Build Cache trên Render dashboard — chỉ chọn nút **"Manual Deploy > Deploy latest commit"** (KHÔNG chọn "Clear build cache & deploy" trừ khi sửa Dockerfile APT).

