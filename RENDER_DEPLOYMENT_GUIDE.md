# CẨM NANG TRIỂN KHAI CLOUD 24/7 LÊN RENDER.COM
## Universal Office & Presentation Studio (Dockerfile Multi-Stage & 512MB RAM Hardening)

> **Mục tiêu**: Đưa toàn bộ hệ sinh thái `antigravity-doc-handler` lên đám mây Render.com hoàn toàn miễn phí, chạy 24/7 trên 1 cổng duy nhất, bạn bè mở link là dùng ngay được đầy đủ tính năng: Chuyển đổi tài liệu, PDF Pro Studio, Tạo bài thuyết trình 16:9, Vẽ sơ đồ kỹ thuật (Mermaid, PlantUML, Draw.io, Canvas).

---

## 🏗️ 1. Kiến Trúc Triển Khai Một Cổng (Unified Single-Port)

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
  - /api/convert          - /assets/*.js, *.css
  - /api/pdf-*            - SPA Fallback Route
  - /api/jobs/* (SSE)
```

- **Stage 1 (Node.js 18)**: Biên dịch toàn bộ mã nguồn Vite frontend (`npm run build`) thành gói tĩnh tối ưu hóa cao tại thư mục `dist/`.
- **Stage 2 (Python 3.11-slim)**: Cài đặt đầy đủ các backend thực thi (LibreOffice, Tesseract OCR, OpenJDK 17, Chromium, Graphviz), sao chép thư mục `dist/` vào container và để Python HTTP Server phục vụ trực tiếp.

---

## 🛡️ 2. Xử Lý Triệt Để 4 Bẫy Thực Chiến (Showstoppers)

| # | Bẫy Tiềm Ẩn | Hậu Quả Nếu Không Xử Lý | Giải Pháp Đã Tích Hợp Sẵn Trong Repo |
|---|---|---|---|
| **1** | **Thiếu Headless Chromium & Graphviz dot** | Lỗi render sơ đồ Draw.io/Mermaid/Canvas; PlantUML báo lỗi *"Cannot find Graphviz dot executable"* khi vẽ Class/C4/UseCase/Activity. | Đã cài `chromium`, `graphviz`, font Noto tiếng Việt trong Dockerfile; cấu hình `ENV CHROMIUM_PATH=/usr/bin/chromium` và cờ `--no-sandbox --disable-dev-shm-usage`. |
| **2** | **Tràn RAM 512MB (OOM Kill - Exit 137)** | Khi nhiều người cùng chuyển đổi tài liệu hoặc chạy song song LibreOffice, Chromium, Java, RAM vượt 512MB làm sập container. | Thiết lập `MAX_JOB_WORKERS=1` để tuần tự hóa hàng đợi công việc; giới hạn heap Java qua `JAVA_TOOL_OPTIONS="-Xmx192m -DPLANTUML_LIMIT_SIZE=16384 -Dfile.encoding=UTF-8"`. |
| **3** | **Bind cứng IP 127.0.0.1 hoặc Port 8000** | Render báo lỗi `Connection Refused` hoặc `Port Binding Timeout` do không thể định tuyến traffic từ ngoài vào container. | `server/app.py` và `web_api_server.py` tự động đọc `HOST=0.0.0.0` và cổng động `PORT=int(os.environ.get("PORT", 8000))`. |
| **4** | **Tràn ổ đĩa tạm (Ephemeral Disk Full)** | Render Free dùng ổ đĩa chia sẻ với bộ nhớ tạm. Các file test PDF/DOCX 50–100MB tích tụ sẽ làm đầy dung lượng. | Tự động chạy `cleanup_old_temp_files(max_age_seconds=7200)` mỗi khi upload để quét sạch các tệp tạm quá 2 giờ. |

---

## 🚀 3. Hướng Dẫn Các Bước Triển Khai Lên Render.com

### Cách 1: Dùng Blueprint `render.yaml` (Khuyến nghị — Nhanh nhất)

1. **Đẩy mã nguồn lên GitHub**:
   ```bash
   git add .
   git commit -m "feat(cloud): production multi-stage dockerfile and render.yaml blueprint"
   git push origin main
   ```
2. **Đăng nhập Render.com**:
   - Truy cập [https://dashboard.render.com/](https://dashboard.render.com/) và đăng nhập bằng tài khoản GitHub của bạn.
3. **Tạo Blueprint**:
   - Bấm nút **New +** ở góc trên bên phải $\rightarrow$ Chọn **Blueprint**.
   - Chọn kho lưu trữ GitHub của bạn (`Offices_Handler` hoặc tên tương đương).
   - Render sẽ tự động nhận diện tệp `render.yaml` có sẵn trong repo và hiển thị cấu hình dịch vụ Web Service.
   - Bấm **Apply**. Quá trình build Docker sẽ tự động bắt đầu.

---

### Cách 2: Tạo Thủ Công (Manual Web Service) Trên Giao Diện Render

Nếu bạn muốn tự thiết lập từng bước:

1. Bấm **New +** $\rightarrow$ Chọn **Web Service**.
2. Chọn repo GitHub của bạn và bấm **Connect**.
3. Điền các trường thông tin cơ bản:
   - **Name**: `offices-handler` (hoặc tên tùy thích).
   - **Region**: `Singapore` (độ trễ thấp nhất cho Việt Nam & Châu Á).
   - **Branch**: `main` (hoặc branch bạn đang sử dụng).
   - **Runtime**: Chọn **Docker**.
   - **Instance Type**: Chọn **Free** (512MB RAM, 0.1 CPU).
4. **Cấu hình biến môi trường (Environment Variables)**:
   Bấm vào mục **Advanced** $\rightarrow$ **Add Environment Variable**:
   - `HOST` = `0.0.0.0`
   - `MAX_JOB_WORKERS` = `1`
   - `JAVA_TOOL_OPTIONS` = `-Xmx192m -DPLANTUML_LIMIT_SIZE=16384 -Dfile.encoding=UTF-8`
   - `CHROMIUM_PATH` = `/usr/bin/chromium`
   - `PYTHONUNBUFFERED` = `1`
   - *(Lưu ý: Không tạo biến `PORT`, hãy để Render tự động gán).*
5. **Cấu hình Health Check Path**:
   - Nhập `/api/health` vào ô **Health Check Path**.
6. Bấm **Create Web Service**.

---

## ⚡ 4. Mẹo Giữ Container Chạy 24/7 Không Bị Sleep (Free Uptime Keep-Alive)

Trên gói Free của Render, container sẽ tự động chuyển sang chế độ ngủ (Sleep) sau 15 phút nếu không có yêu cầu truy cập nào. Khi có người mở link, container sẽ mất khoảng 30–50 giây để khởi động lại (Cold Start).

Để container luôn thức 24/7 và sẵn sàng phục vụ tức thì:

1. Truy cập một dịch vụ giám sát miễn phí (chọn 1 trong 2):
   - **cron-job.org** ([https://cron-job.org/](https://cron-job.org/))
   - **UptimeRobot** ([https://uptimerobot.com/](https://uptimerobot.com/))
2. Tạo một tác vụ Ping định kỳ:
   - **URL cần ping**: `https://<ten-ung-dung-cua-ban>.onrender.com/api/health`
   - **Chu kỳ**: Mỗi `10 phút` (hoặc `14 phút`) 1 lần.
   - **Phương thức**: `GET`
3. Endpoint `/api/health` siêu nhẹ, phản hồi JSON chỉ mất vài mili-giây và không tốn tài nguyên CPU, giúp container luôn ở trạng thái `online` 24/7.

---

## 🔍 5. Kiểm Tra Sau Khi Triển Khai (Verification Checklist)

Khi Render báo trạng thái **Live**:
- [x] Mở trang chủ `https://<app>.onrender.com/`: Giao diện Web Studio hiện lên nhanh chóng với font chữ Outfit/Inter chuẩn đẹp.
- [x] Kiểm tra chẩn đoán: Vào endpoint `/api/doctor` hoặc bấm biểu tượng Doctor trên thanh điều hướng để xác nhận Java, PlantUML, Chromium, LibreOffice, Tesseract đều hiển thị trạng thái `healthy`.
- [x] Chuyển đổi thử một tệp PDF/DOCX để kiểm tra hàng đợi SSE Job Manager phản hồi thanh tiến trình mượt mà.
- [x] Mở tab Diagram Studio và vẽ thử một sơ đồ (Mermaid hoặc Draw.io) để xác nhận xuất file ảnh PNG 300+ DPI thành công.
