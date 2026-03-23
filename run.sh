#!/bin/bash

echo "==============================================="
echo "   Cdimex Mail Merge - Khởi động ứng dụng"
echo "==============================================="
echo

# Kiểm tra virtual environment
if [ ! -d ".venv" ]; then
    echo "[ERROR] Môi trường ảo chưa được tạo!"
    echo "Vui lòng chạy file 'install.sh' trước."
    exit 1
fi

# Kích hoạt virtual environment và chạy ứng dụng
echo "[INFO] Khởi động ứng dụng..."
source .venv/bin/activate
streamlit run streamlit_app.py

echo
echo "[INFO] Ứng dụng đã dừng."


