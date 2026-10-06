# Hướng Dẫn Sử Dụng & Khởi Chạy Excel Data Entry Web Trên macOS

Bản web **Excel Data Entry Web** đã được build độc lập (`self-contained`), sẵn sàng chạy trực tiếp trên macOS mà **KHÔNG cần cài đặt .NET SDK hay Runtime**.

Hỗ trợ 2 kiến trúc máy Mac:
- **`publish/osx-arm64`**: Dành cho Mac chạy chip Apple Silicon (**M1, M2, M3, M4, Mac Studio, v.v.**)
- **`publish/osx-x64`**: Dành cho Mac chạy chip **Intel**.

---

## 🌟 CÁCH 1: Mở trực tiếp bằng cách nhấp đúp chuột (Khuyên dùng - Không cần gõ Terminal)

Mở Finder vào thư mục tương ứng với máy của bạn (`publish/osx-arm64` hoặc `publish/osx-x64`):

### 1. Khởi chạy nhanh (Chạy ngầm + Tự mở trình duyệt):
- **Nhấp đúp chuột vào file `start.command`**.
- Ứng dụng sẽ tự động chạy ngầm và Safari / Chrome sẽ tự động bật lên mở trang **`http://localhost:5000`**.
- Khi muốn tắt ứng dụng: Nhấp đúp chuột vào **`stop.command`**.

### 2. Tạo phím tắt mở trực tiếp trên Desktop (Màn hình chính macOS):
1. Nhấp đúp vào file **`create_desktop_shortcut.command`**.
2. Trên màn hình Desktop sẽ xuất hiện file **`ExcelDataEntryWeb.command`**.
3. Từ lần sau, bạn chỉ cần nhấp đúp vào biểu tượng trên Desktop là ứng dụng tự động mở!

### 3. Khởi chạy có theo dõi Log (Chạy trực tiếp trong Terminal):
- Nhấp đúp vào file **`run.command`**.
- Cửa sổ Terminal sẽ hiển thị trạng thái và log chi tiết. Nhấn `Ctrl + C` để dừng.

---

## 💡 CÁCH 2: Chạy qua Terminal (Dành cho Developer / Power User)

Mở **Terminal**, chuyển vào thư mục bản build và gõ:

```bash
# Cấp quyền thực thi và gỡ Gatekeeper quarantine (chỉ cần chạy lần đầu nếu tải từ mạng về)
chmod +x ./ExcelDataEntryWeb ./*.command ./*.sh
xattr -d com.apple.quarantine ./ExcelDataEntryWeb 2>/dev/null || true

# Khởi chạy
./start.command
```

---

## ⚠️ Lưu ý bảo mật trên macOS (Gatekeeper):
Nếu lần đầu tiên mở file `.command` hoặc file thực thi mà macOS hiện thông báo *"App cannot be opened because it is from an unidentified developer"* hoặc *"App is damaged"*:
1. **Cách 1**: Giữ phím `Control` (hoặc chuột phải) trên file -> chọn **Open** -> chọn **Open** một lần nữa.
2. **Cách 2**: Mở **System Settings** (Cài đặt hệ thống) -> **Privacy & Security** (Bảo mật & Quyền riêng tư) -> cuộn xuống tìm mục Security và bấm **Open Anyway** (Vẫn mở).
3. **Cách 3**: Mở Terminal tại thư mục ứng dụng và gõ:
   ```bash
   xattr -cr .
   chmod +x ./ExcelDataEntryWeb ./*.command ./*.sh
   ```

---

## 🚀 Thao tác nhập liệu Excel trên macOS:
- Tương thích 100% với Safari, Google Chrome, Microsoft Edge, Arc trên macOS.
- Hỗ trợ đầy đủ tải file Excel lên, xem trước bảng tính, nhập liệu nhanh và lưu/tải về file Excel đã cập nhật.
