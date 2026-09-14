# Hướng Dẫn Deploy Bản Web Trực Tiếp Trên Linux

Bản web **Excel Data Entry Web** đã được chuẩn hóa đa nền tảng (Cross-platform), tương thích 100% với các hệ điều hành Linux (Ubuntu, Debian, CentOS, Fedora, Arch Linux,...).

---

## CÁCH 1: Nhấn đúp chuột để mở trực tiếp (Khuyên dùng - Không cần gõ Terminal)

Trong thư mục `publish/linux-x64` đã có sẵn các file tiện ích:

### 🌟 Lựa chọn A: Mở ngay bằng file `start.sh`
- Bạn chỉ cần **nhấn đúp chuột vào file `start.sh`** (hoặc chọn "Run / Run in Terminal" trên giao diện Linux).
- Script sẽ **tự động khởi động ứng dụng chạy ngầm** và **tự động bật trình duyệt web** (Chrome / Firefox) mở thẳng trang `http://localhost:5000`.
- Khi không dùng nữa, nhấn đúp vào `stop.sh` để tắt ứng dụng.

### 🌟 Lựa chọn B: Tạo biểu tượng mở trực tiếp trên Desktop (Màn hình chính)
Nếu bạn muốn có một biểu tượng ứng dụng có logo Excel nằm ngoài màn hình Desktop:
1. Nhấp đúp (hoặc chạy 1 lần duy nhất) file:
   ```bash
   ./create_desktop_shortcut.sh
   ```
2. Trên màn hình Desktop máy Linux sẽ xuất hiện ngay biểu tượng **"Excel Data Entry Web"** với logo xanh chuyên nghiệp.
3. Từ lần sau, bạn chỉ việc **nhấp đúp vào icon trên Desktop** là ứng dụng tự mở trình duyệt lên nhập liệu ngay, hoàn toàn không cần mở Terminal!

---

## CÁCH 2: Chạy trực tiếp qua Terminal (Nếu muốn xem log chi tiết)
> **Ưu điểm**: File chạy độc lập (`self-contained`), **KHÔNG cần cài .NET SDK/Runtime trên máy Linux**.

Mở Terminal tại thư mục chứa file và gõ:
```bash
chmod +x ./ExcelDataEntryWeb ./start.sh ./run.sh
./start.sh
```

---

## CÁCH 2: Chạy bằng Docker / Docker Compose

Nếu máy Linux đã cài đặt Docker:

### Cách chạy với Dockerfile:
```bash
# 1. Build image:
docker build -t excel-web .

# 2. Chạy container:
docker run -d -p 5000:5000 --name excel-web-app --restart unless-stopped excel-web
```

### Hoặc với Docker Compose:
```bash
docker compose up -d
```

Sau đó mở trình duyệt truy cập: **`http://localhost:5000`**

---

## CÁCH 3: Cài đặt dịch vụ chạy ngầm tự khởi động (systemd Service)

Để ứng dụng tự động chạy ngầm và tự khởi động lại khi bật máy Linux:

1. Copy thư mục `publish/linux-x64` vào `/opt/excel-web`:
   ```bash
   sudo mkdir -p /opt/excel-web
   sudo cp -r publish/linux-x64/* /opt/excel-web/
   sudo chmod +x /opt/excel-web/ExcelDataEntryWeb
   ```

2. Tạo file cấu hình service:
   ```bash
   sudo nano /etc/systemd/system/excel-web.service
   ```

3. Dán nội dung sau vào:
   ```ini
   [Unit]
   Description=Excel Data Entry Web Service
   After=network.target

   [Service]
   WorkingDirectory=/opt/excel-web
   ExecStart=/opt/excel-web/ExcelDataEntryWeb --urls "http://0.0.0.0:5000"
   Restart=always
   RestartSec=10
   SyslogIdentifier=excel-web
   User=root
   Environment=ASPNETCORE_ENVIRONMENT=Production

   [Install]
   WantedBy=multi-user.target
   ```

4. Kích hoạt và khởi động:
   ```bash
   sudo systemctl daemon-reload
   sudo systemctl enable excel-web
   sudo systemctl start excel-web
   ```

5. Kiểm tra trạng thái:
   ```bash
   sudo systemctl status excel-web
   ```

---

## Lưu ý khi thao tác file Excel trên Linux bằng trình duyệt:
- **Chọn file / Tải file**: Trình duyệt trên Linux (Chrome / Firefox) hỗ trợ tải lên file Excel từ máy Linux và lưu lại / tải về hoàn toàn tương tự như trên Windows.
- **Trình duyệt khuyến nghị**: Google Chrome, Chromium hoặc Firefox phiên bản mới để tận dụng tốt nhất File System Access API.
