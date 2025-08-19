# Bookmedi Mail Merge

Ứng dụng gửi email hàng loạt kèm file PDF riêng cho từng người nhận.

## Cài đặt nhanh (cho người không biết lập trình)

### Bước 1: Tải và giải nén
1. Tải toàn bộ thư mục này về máy
2. Giải nén vào một thư mục (ví dụ: Desktop)

### Bước 2: Chạy cài đặt tự động
1. Mở Terminal (macOS/Linux) hoặc Command Prompt (Windows)
2. Di chuyển vào thư mục đã giải nén:
   ```bash
   cd đường/dẫn/tới/thư/mục
   ```
3. Chạy script cài đặt:
   ```bash
   python install.py
   ```
4. Làm theo hướng dẫn trên màn hình

### Bước 3: Chạy ứng dụng
- **Windows**: Double-click file `run_gui.bat`
- **macOS/Linux**: Double-click file `run_gui.sh` hoặc chạy `./run_gui.sh`

Trình duyệt sẽ tự động mở tại `http://localhost:8501`

## Cách sử dụng

### 1. Chuẩn bị file
- **recipients.xlsx**: Danh sách người nhận (Email, Ten, FilePDF, Subject, CC, BCC)
- **template.html**: Mẫu email HTML
- **PDF files**: Các file PDF cần đính kèm

### 2. Cấu hình SMTP
Trong giao diện web, điền thông tin SMTP:
- **Gmail**: smtp.gmail.com, port 587, dùng App Password
- **Outlook**: smtp.office365.com, port 587, dùng App Password

### 3. Gửi email
1. Upload file recipients.xlsx và template.html
2. Điền thông tin SMTP
3. Bật "Dry-run" để thử trước
4. Bấm "Gửi Email"

## Cấu trúc file recipients.xlsx

| Cột | Bắt buộc | Mô tả |
|-----|----------|-------|
| Email | ✓ | Email người nhận |
| Ten | ✓ | Tên người nhận (dùng trong template) |
| FilePDF | ✓ | Đường dẫn file PDF hoặc URL |
| Subject | | Tiêu đề email (có thể dùng {{Ten}}) |
| CC | | Email CC (phân cách bằng dấu phẩy) |
| BCC | | Email BCC (phân cách bằng dấu phẩy) |

## Template HTML

Sử dụng `{{Ten}}`, `{{Email}}`, `{{NgayGui}}` trong template để thay thế động.

Ví dụ:
```html
<p>Dear bạn {{Ten}},</p>
<p>Bookmedi gửi bạn kết quả bài thi Versant level 1.</p>
```

## Cách sử dụng nâng cao

### Gửi bằng Command Line
```bash
python send_mail_merge.py \
  --recipients recipients.xlsx \
  --template template.html \
  --smtp-host smtp.gmail.com \
  --smtp-port 587 \
  --smtp-user your-email@gmail.com \
  --smtp-pass your-app-password \
  --from-name "Bookmedi" \
  --dry-run
```

### Deploy lên server
1. Upload toàn bộ code lên server
2. Chạy `python install.py` trên server
3. Chạy `streamlit run streamlit_app.py --server.port 8501 --server.address 0.0.0.0`
4. Truy cập qua IP server:port

## Xử lý lỗi thường gặp

### Lỗi "No module named pandas"
- Chạy lại `python install.py`
- Đảm bảo đang dùng Python 3.8+

### Lỗi SMTP
- Kiểm tra email/password
- Gmail: Bật 2FA và tạo App Password
- Outlook: Bật SMTP AUTH

### File PDF không tìm thấy
- Kiểm tra đường dẫn trong cột FilePDF
- Dùng đường dẫn tuyệt đối hoặc URL
- Điền Base directory trong giao diện web

## Hỗ trợ

Nếu gặp vấn đề, hãy:
1. Kiểm tra log trong giao diện web
2. Chạy với `--dry-run` trước
3. Liên hệ admin để được hỗ trợ