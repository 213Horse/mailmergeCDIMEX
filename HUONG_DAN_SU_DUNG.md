# Cdimex Mail Merge - Hướng dẫn sử dụng

## 🚀 Cài đặt nhanh (1 lệnh)

### Trên Windows:
1. Double-click file `install.bat`
2. Đợi cài đặt hoàn tất
3. Double-click file `run.bat` để chạy ứng dụng

### Trên macOS/Linux:
1. Mở Terminal, chạy: `./install.sh`
2. Đợi cài đặt hoàn tất  
3. Chạy: `./run.sh` để khởi động ứng dụng

## 📋 Yêu cầu hệ thống

- **Python 3.8+** (sẽ được kiểm tra tự động)
- **Windows 10+** hoặc **macOS 10.14+** hoặc **Linux**
- **Kết nối Internet** (để tải thư viện)

## 🎯 Chức năng chính

### 1. Gửi Email Hàng Loạt
- Upload file Excel/CSV chứa danh sách khách hàng
- Soạn nội dung email trực tiếp hoặc dùng file HTML
- Gửi kèm file PDF cho từng khách hàng
- Hỗ trợ token động: `{{Ten}}`, `{{Email}}`, `{{NgayGui}}`

### 2. Quản lý File
- Upload và quản lý file trong thư mục dự án
- Tạo thư mục mới
- Xem, tải xuống, xóa file

## 📁 Cấu trúc file quan trọng

```
mailmerge/
├── install.bat          # Cài đặt trên Windows
├── install.sh           # Cài đặt trên macOS/Linux  
├── run.bat              # Chạy ứng dụng trên Windows
├── run.sh               # Chạy ứng dụng trên macOS/Linux
├── recipients.xlsx      # File danh sách khách hàng mẫu
├── template.html        # Template email mẫu
├── header.html          # Header cố định cho email
├── footer.html          # Footer cố định cho email
└── uploads/             # Thư mục chứa file PDF, ảnh
```

## ⚙️ Cấu hình SMTP

### Gmail:
- **SMTP Host:** smtp.gmail.com
- **Port:** 587
- **User:** email@gmail.com
- **Password:** App Password (không phải mật khẩu thường)

### Outlook:
- **SMTP Host:** smtp.office365.com  
- **Port:** 587
- **User:** email@outlook.com
- **Password:** mật khẩu email

## 🔧 Xử lý sự cố

### Lỗi "Python chưa được cài đặt":
- **Windows:** Tải từ https://python.org
- **macOS:** `brew install python3`
- **Linux:** `sudo apt install python3 python3-pip python3-venv`

### Lỗi "Không thể cài đặt thư viện":
- Kiểm tra kết nối Internet
- Chạy lại `install.bat` hoặc `install.sh`

### Ứng dụng không mở:
- Kiểm tra file `.venv` có tồn tại không
- Chạy lại `install.bat` hoặc `install.sh`

### Lỗi gửi email:
- Kiểm tra thông tin SMTP
- Bật "Dry-run" để test trước
- Kiểm tra App Password cho Gmail

## 📞 Hỗ trợ

Nếu gặp vấn đề, vui lòng:
1. Kiểm tra file log trong ứng dụng
2. Chạy ở chế độ "Dry-run" trước
3. Liên hệ team phát triển

---
**Lưu ý:** Giữ cửa sổ Terminal/Command Prompt mở khi sử dụng ứng dụng.


