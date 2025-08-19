#!/usr/bin/env bash
set -Eeuo pipefail

# Deployment script for Bookmedi Mail Merge (runs on remote server)
# - Creates/updates venv
# - Installs requirements
# - Restarts Streamlit app

APP_DIR="$(cd "$(dirname "$0")/.." && pwd)"
PORT="${STREAMLIT_PORT:-8502}"           # khớp default với workflow (8502)
ADDRESS="${STREAMLIT_ADDRESS:-0.0.0.0}"
APP_FILE="${APP_FILE:-streamlit_app.py}" # đổi nếu entrypoint khác
PY_BIN="python3"
command -v "${PY_BIN}" >/dev/null 2>&1 || PY_BIN="python"

cd "$APP_DIR"

# Create venv if missing
if [[ ! -d .venv ]]; then
  echo "[deploy] Creating virtual env..."
  "${PY_BIN}" -m venv .venv
fi

# Activate venv
# shellcheck disable=SC1091
source .venv/bin/activate

# (tuỳ chọn) mirror pip như log của bạn
: "${PIP_INDEX_URL:=https://mirrors.aliyun.com/pypi/simple}"
export PIP_INDEX_URL

# Upgrade pip và cài deps
python -m pip install --upgrade pip
if [[ -f requirements.txt ]]; then
  echo "[deploy] Installing requirements..."
  pip install -r requirements.txt
else
  echo "[deploy] requirements.txt not found; installing minimal deps..."
  pip install pandas openpyxl streamlit requests
fi

# Stop tiến trình đang chiếm port (best effort)
echo "[deploy] Stopping existing app on port ${PORT} (if any)..."
if command -v fuser >/dev/null 2>&1; then
  fuser -k "${PORT}/tcp" || true
else
  pids="$(lsof -ti tcp:${PORT} || true)"
  if [[ -n "${pids}" ]]; then
    kill -9 ${pids} || true
  fi
fi

# Start app
echo "[deploy] Starting Streamlit app on ${ADDRESS}:${PORT}..."
mkdir -p logs
STREAMLIT_BIN="$(pwd)/.venv/bin/streamlit"

nohup "${STREAMLIT_BIN}" run "${APP_FILE}" \
  --server.address "${ADDRESS}" \
  --server.port "${PORT}" \
  --server.headless true \
  > "logs/streamlit.log" 2>&1 &

sleep 1
echo "[deploy] Done. Logs: ${APP_DIR}/logs/streamlit.log"
