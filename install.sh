#!/bin/bash

echo "==============================================="
echo "   Cdimex Mail Merge - Cài đặt tự động"
echo "==============================================="
echo

# Kiểm tra Python
if ! command -v python3 &> /dev/null; then
    echo "[ERROR] Python3 chưa được cài đặt!"
    echo "Vui lòng cài đặt Python3 trước:"
    echo "  - macOS: brew install python3"
    echo "  - Ubuntu/Debian: sudo apt install python3 python3-pip python3-venv"
    echo "  - CentOS/RHEL: sudo yum install python3 python3-pip"
    exit 1
fi

echo "[OK] Python3 đã được cài đặt"
echo

# Tạo virtual environment
echo "[INFO] Tạo môi trường ảo..."
python3 -m venv .venv
if [ $? -ne 0 ]; then
    echo "[ERROR] Không thể tạo môi trường ảo"
    exit 1
fi

# Kích hoạt virtual environment và cài đặt packages
echo "[INFO] Cài đặt các thư viện cần thiết..."
source .venv/bin/activate
python -m pip install --upgrade pip
pip install pandas openpyxl streamlit requests streamlit-quill

if [ $? -ne 0 ]; then
    echo "[ERROR] Không thể cài đặt thư viện"
    exit 1
fi

echo
echo "==============================================="
echo "   Cài đặt hoàn tất!"
echo "==============================================="
echo
echo "Để chạy ứng dụng:"
echo "1. Double-click file 'run.sh'"
echo "2. Hoặc chạy lệnh: ./run.sh"
echo


