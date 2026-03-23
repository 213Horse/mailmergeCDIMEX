# 📧 Cdimex Mail Merge

Ứng dụng gửi email hàng loạt kèm file PDF cho từng khách hàng.

## 🚀 Cài đặt và chạy (1 lệnh)

### Windows:
1. **Cài đặt:** Double-click `install.bat`
2. **Chạy:** Double-click `run.bat`

### macOS/Linux:
1. **Cài đặt:** `./install.sh`
2. **Chạy:** `./run.sh`

## 📋 Yêu cầu

- Python 3.8+ (sẽ được kiểm tra tự động)
- Kết nối Internet (để tải thư viện)

## 🎯 Chức năng

- ✅ Gửi email hàng loạt từ file Excel/CSV
- ✅ Soạn nội dung email trực tiếp (WYSIWYG)
- ✅ Gửi kèm file PDF riêng cho từng khách hàng
- ✅ Hỗ trợ token động: `{{Ten}}`, `{{Email}}`, `{{NgayGui}}`
- ✅ Quản lý file và thư mục
- ✅ Giao diện web đẹp và dễ sử dụng

## 📁 File quan trọng

| File | Mô tả |
|------|-------|
| `install.bat` / `install.sh` | Cài đặt tự động |
| `run.bat` / `run.sh` | Chạy ứng dụng |
| `check_system.bat` / `check_system.sh` | Kiểm tra hệ thống |
| `recipients.xlsx` | File danh sách khách hàng mẫu |
| `template.html` | Template email mẫu |
| `uploads/` | Thư mục chứa file PDF, ảnh |

## ⚙️ Cấu hình SMTP

### Gmail:
- Host: `smtp.gmail.com`
- Port: `587`
- Password: App Password (không phải mật khẩu thường)

### Outlook:
- Host: `smtp.office365.com`
- Port: `587`

## 🔧 Xử lý sự cố

### Lỗi cài đặt:
- Chạy `check_system.bat` / `check_system.sh` để kiểm tra
- Đảm bảo có kết nối Internet
- Cài đặt Python từ https://python.org

### Lỗi gửi email:
- Kiểm tra thông tin SMTP
- Bật "Dry-run" để test trước
- Kiểm tra App Password cho Gmail

## 📖 Hướng dẫn chi tiết

- **Hướng dẫn nhanh:** `QUICK_START.md`
- **Hướng dẫn đầy đủ:** `HUONG_DAN_SU_DUNG.md`

## ⚠️ Lưu ý

- **Giữ cửa sổ Terminal/Command Prompt mở** khi sử dụng
- **Test trước** bằng chế độ "Dry-run"
- **Backup dữ liệu** trước khi gửi hàng loạt

---
**Phát triển bởi:** Cdimex Team