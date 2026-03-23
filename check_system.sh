#!/bin/bash

echo "==============================================="
echo "   Cdimex Mail Merge - Kiểm tra hệ thống"
echo "==============================================="
echo

echo "[INFO] Kiểm tra Python..."
if command -v python3 &> /dev/null; then
    python3 --version
    echo "[OK] Python đã sẵn sàng"
else
    echo "[ERROR] Python3 chưa được cài đặt!"
    echo "Vui lòng cài đặt Python3 trước"
    exit 1
fi
echo

echo "[INFO] Kiểm tra virtual environment..."
if [ -d ".venv" ]; then
    echo "[OK] Virtual environment đã tồn tại"
    source .venv/bin/activate
    echo "[INFO] Kiểm tra các thư viện..."
    python -c "import pandas, openpyxl, streamlit, requests; print('[OK] Tất cả thư viện đã sẵn sàng')" 2>/dev/null
    if [ $? -ne 0 ]; then
        echo "[WARNING] Một số thư viện chưa được cài đặt"
        echo "Chạy ./install.sh để cài đặt lại"
    fi
else
    echo "[WARNING] Virtual environment chưa được tạo"
    echo "Chạy ./install.sh để cài đặt"
fi
echo

echo "[INFO] Kiểm tra file cần thiết..."
if [ -f "streamlit_app.py" ]; then
    echo "[OK] streamlit_app.py"
else
    echo "[ERROR] Không tìm thấy streamlit_app.py"
fi

if [ -f "recipients.xlsx" ]; then
    echo "[OK] recipients.xlsx"
else
    echo "[WARNING] Không tìm thấy recipients.xlsx (có thể tạo mới)"
fi

if [ -f "template.html" ]; then
    echo "[OK] template.html"
else
    echo "[WARNING] Không tìm thấy template.html (có thể tạo mới)"
fi

echo
echo "==============================================="
echo "   Kiểm tra hoàn tất"
echo "==============================================="
echo


