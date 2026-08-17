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

# Nạp biến môi trường nếu có file .env (tương thích cả local & server)
if [ -f ".env" ]; then
  set -o allexport
  # shellcheck disable=SC1091
  source ".env" || true
  set +o allexport
fi

PORT="${STREAMLIT_SERVER_PORT:-6520}"
echo "[INFO] Streamlit sẽ chạy tại: http://0.0.0.0:${PORT}"
streamlit run streamlit_app.py \
  --server.address="0.0.0.0" \
  --server.port="${PORT}" \
  --server.headless="true" \
  --browser.gatherUsageStats="false"

echo
echo "[INFO] Ứng dụng đã dừng."


