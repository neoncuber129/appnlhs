# 📊 Excel Data Entry Web (APP Nhập Liệu Hồ Sơ Sức Khỏe & Excel)

[![Build & Deploy](https://github.com/neoncuber129/appnlhs/actions/workflows/deploy.yml/badge.svg)](https://github.com/neoncuber129/appnlhs/actions/workflows/deploy.yml)
[![Docker Image](https://img.shields.io/badge/Docker-GHCR-blue?logo=docker)](https://github.com/neoncuber129/appnlhs/pkgs/container/appnlhs)
[![.NET 10](https://img.shields.io/badge/.NET-10.0-purple?logo=dotnet)](https://dotnet.microsoft.com/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

> **Ứng dụng Web chuyên nghiệp dành cho nhập liệu, chuẩn hóa dữ liệu Excel, quản lý hồ sơ khám sức khỏe theo Thông tư 37 (TT37) và xuất báo cáo Word.**

---

## 🚀 Triển Khai Nhanh 1-Click Lên Web Trực Tuyến

Bạn có thể đưa toàn bộ ứng dụng này lên mạng Internet miễn phí chỉ với **1 cú nhấp chuột**:

[![Deploy to Render](https://render.com/images/deploy-to-render-button.svg)](https://render.com/deploy?repo=https://github.com/neoncuber129/appnlhs)

---

## ✨ Các Tính Năng Nổi Bật

- ⚡ **Nhập liệu siêu tốc**: Tối ưu phím tắt (Enter, Tab, mũi tên, Shift+Enter), tự động nhảy dòng, nhảy cột thông minh.
- 🎨 **Giao diện hiện đại**: Thiết kế Web chuẩn hiện đại, hỗ trợ Dark / Light Theme, biểu đồ thống kê trực quan.
- 📑 **Quản lý Sheet linh hoạt**: Chọn Sheet, ẩn/hiện cột theo ý muốn, tùy chỉnh dòng tiêu đề và dòng dữ liệu mẫu.
- 🔍 **Tìm kiếm & Lọc dữ liệu**: Tìm kiếm nhanh mọi trường thông tin, kiểm tra trùng lặp, lọc theo phòng khám/chuyên khoa.
- 🖨️ **Xuất phiếu & In ấn TT37**: Xuất file Word (docx) phiếu khám sức khỏe SQ, QNCN theo mẫu chuẩn Thông tư 37.
- 🐳 **Đa nền tảng**: Chạy tốt trên Web Browser, Docker, Linux (Ubuntu/Debian/CentOS), Windows và MacOS.

---

## 🛠️ Hướng Dẫn Sử Dụng & Triển Khai

### 1. Chạy nhanh bằng Docker (Khuyên dùng)

#### Cách A: Chạy trực tiếp qua Docker Container từ GitHub Registry
```bash
docker run -d -p 5000:5000 --name appnlhs-web --restart unless-stopped ghcr.io/neoncuber129/appnlhs:latest
```
Sau đó mở trình duyệt truy cập: **`http://localhost:5000`**

#### Cách B: Sử dụng Docker Compose
```bash
git clone https://github.com/neoncuber129/appnlhs.git
cd appnlhs
docker compose up -d
```

---

### 2. Chạy từ Mã Nguồn (.NET 10 SDK)

Yêu cầu đã cài đặt [.NET 10.0 SDK](https://dotnet.microsoft.com/download):

```bash
# 1. Clone repository
git clone https://github.com/neoncuber129/appnlhs.git
cd appnlhs

# 2. Khởi chạy Web Server
dotnet run --project ExcelDataEntryWeb/ExcelDataEntryWeb.csproj
```
Trình duyệt tự động mở hoặc truy cập: **`http://localhost:5000`**

---

### 3. Chạy Bản Độc Lập Không Cần Cài .NET (Standalone Binaries)

Tải gói phát hành trong mục [Releases / Actions Artifacts](https://github.com/neoncuber129/appnlhs/actions):
- **Linux**: Giải nén và chạy `./start.sh` hoặc `./ExcelDataEntryWeb`.
- **Windows**: Giải nén và chạy `ExcelDataEntryWeb.exe`.

Chi tiết xem thêm tại [DEPLOY_LINUX.md](DEPLOY_LINUX.md).

---

## 📁 Cấu Trúc Dự Án

```
├── .github/workflows/         # CI/CD Tự động hóa build Docker & Web Binaries
├── Dockerfile                 # Đóng gói container đa nền tảng
├── docker-compose.yml         # File chạy Docker Compose
├── render.yaml                # Cấu hình 1-Click Deploy lên Cloud (Render)
├── ExcelDataEntryCore/        # Thư viện xử lý logic dữ liệu Excel & Word (OpenXML/EPPlus)
├── ExcelDataEntryWeb/         # Ứng dụng Web Server ASP.NET Core & Giao diện frontend
│   └── wwwroot/               # Giao diện HTML5, CSS3, JavaScript tương tác người dùng
└── MAU 01- TT37...docx        # Mẫu phiếu khám sức khỏe chuẩn TT37
```

---

## 📜 Giấy Phép (License)

Dự án được phát hành dưới giấy phép [MIT License](LICENSE).
