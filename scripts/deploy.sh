#!/usr/bin/env bash
set -euo pipefail

# Deployment script for Bookmedi Mail Merge (runs on remote server)
# - Creates/updates venv
# - Installs requirements
# - Restarts Streamlit app

APP_DIR="$(cd "$(dirname "$0")/.." && pwd)"
PORT="${STREAMLIT_PORT:-8501}"
ADDRESS="${STREAMLIT_ADDRESS:-0.0.0.0}"
PY_BIN="python3"

if ! command -v "${PY_BIN}" >/dev/null 2>&1; then
  PY_BIN="python"
fi

cd "$APP_DIR"

# Create venv if missing
if [[ ! -d .venv ]]; then
  echo "[deploy] Creating virtual env..."
  "${PY_BIN}" -m venv .venv
fi

# Activate venv
# shellcheck disable=SC1091
source .venv/bin/activate

# Upgrade pip and install deps
python -m pip install --upgrade pip
if [[ -f requirements.txt ]]; then
  echo "[deploy] Installing requirements..."
  pip install -r requirements.txt
else
  echo "[deploy] requirements.txt not found; installing minimal deps..."
  pip install pandas openpyxl streamlit requests
fi

# Stop existing Streamlit process on the port (best effort)
echo "[deploy] Stopping existing app on port ${PORT} (if any)..."
if command -v lsof >/dev/null 2>&1; then
  pids=$(lsof -ti tcp:"${PORT}" || true)
  if [[ -n "${pids}" ]]; then
    echo "[deploy] Killing PIDs: ${pids}"
    kill -9 ${pids} || true
  fi
else
  pkill -f "streamlit run streamlit_app.py" || true
fi

# Start the app
echo "[deploy] Starting Streamlit app on ${ADDRESS}:${PORT}..."
nohup .venv/bin/python -m streamlit run streamlit_app.py \
  --server.address "${ADDRESS}" \
  --server.port "${PORT}" \
  > app.log 2>&1 &

echo "[deploy] Done. Logs: ${APP_DIR}/app.log
