#!/usr/bin/env bash
set -euo pipefail

# Chuẩn hóa tên container/app
SAFE_IMAGE_NAME=$(echo "${IMAGE_NAME:-app}" | tr '[:upper:]' '[:lower:]' | sed -E 's/[^a-z0-9._-]+/-/g')
APP_DIR="/opt/${SAFE_IMAGE_NAME}"
HOST_PORT_RAW="${HOST_PORT:-}"
CONTAINER_PORT_RAW="${CONTAINER_PORT:-}"
HOST_PORT="${HOST_PORT_RAW:-6520}"
CONTAINER_PORT="${CONTAINER_PORT_RAW:-6520}"

echo "🚀 Deploying $SAFE_IMAGE_NAME (image: ${REPO_PATH}:${TAG_SHA})"

# Tạo thư mục ứng dụng
sudo mkdir -p "$APP_DIR"

# Ghi file .env
# Lưu ý: truyền PROD_ENV_FILE trực tiếp qua SSH dễ vỡ vì multiline.
# Ưu tiên dùng PROD_ENV_FILE_B64 (base64, 1 dòng) và decode trên server.
if [ -n "${PROD_ENV_FILE_B64:-}" ]; then
  printf "%s" "${PROD_ENV_FILE_B64}" | base64 -d | sudo tee "$APP_DIR/.env" > /dev/null
elif [ -n "${PROD_ENV_FILE:-}" ]; then
  printf "%s" "${PROD_ENV_FILE}" | sudo tee "$APP_DIR/.env" > /dev/null
else
  sudo touch "$APP_DIR/.env"
fi

# Nếu không truyền CONTAINER_PORT, ưu tiên đọc từ .env (STREAMLIT_SERVER_PORT)
if [ -z "${CONTAINER_PORT_RAW}" ]; then
  ENV_STREAMLIT_PORT="$(sudo awk -F= '/^STREAMLIT_SERVER_PORT=/{p=$2} END{print p}' "$APP_DIR/.env" | tr -d '\r' | tr -d ' ')"
  if [ -n "${ENV_STREAMLIT_PORT}" ]; then
    CONTAINER_PORT="${ENV_STREAMLIT_PORT}"
  fi
fi

# Đảm bảo docker compose plugin có sẵn
DOCKER_COMPOSE_CMD=""
if docker compose version >/dev/null 2>&1; then
  DOCKER_COMPOSE_CMD="docker compose"
elif docker-compose version >/dev/null 2>&1; then
  DOCKER_COMPOSE_CMD="docker-compose"
else
  if [ -f /etc/os-release ]; then
    . /etc/os-release
    if [ "${ID}" = "ubuntu" ] || [ "${ID}" = "debian" ]; then
      sudo apt-get update -y
      # Try v2 plugin first (may require Docker's apt repo on older distros)
      if sudo apt-get install -y docker-compose-plugin; then
        DOCKER_COMPOSE_CMD="docker compose"
      else
        # Fallback: v1 docker-compose package (available on many older distros)
        sudo apt-get update -y
        sudo apt-get install -y docker-compose
        DOCKER_COMPOSE_CMD="docker-compose"
      fi
    fi
  fi
fi

if [ -z "${DOCKER_COMPOSE_CMD}" ]; then
  echo "[ERROR] Cannot find or install Docker Compose (v2 plugin or v1 docker-compose)."
  exit 1
fi

# Login GHCR để pull image private
echo "${GHCR_TOKEN}" | sudo docker login ghcr.io -u "${GHCR_USERNAME}" --password-stdin

# Ghi file docker-compose.yml
sudo tee "$APP_DIR/docker-compose.yml" > /dev/null <<YAML
services:
  app:
    image: ${REPO_PATH}:${TAG_SHA}
    container_name: ${SAFE_IMAGE_NAME}
    restart: unless-stopped
    env_file:
      - .env
    ports:
      - "0.0.0.0:${HOST_PORT}:${CONTAINER_PORT}"
YAML

# Pull & chạy container
export COMPOSE_PROJECT_NAME="${SAFE_IMAGE_NAME}"
sudo -E ${DOCKER_COMPOSE_CMD} -f "$APP_DIR/docker-compose.yml" pull app
sudo -E ${DOCKER_COMPOSE_CMD} -f "$APP_DIR/docker-compose.yml" up -d --remove-orphans

# Post-deploy quick diagnostics (giúp bắt lỗi sai port/crash/firewall)
echo "🔎 Checking container status..."
sudo docker ps --filter "name=^/${SAFE_IMAGE_NAME}$" --format "table {{.Names}}\t{{.Status}}\t{{.Ports}}" || true

echo "🔎 Checking local HTTP on VPS (host port: ${HOST_PORT})..."
if command -v curl >/dev/null 2>&1; then
  # Streamlit health endpoint (newer versions) may exist; fallback to /
  (curl -fsS "http://127.0.0.1:${HOST_PORT}/_stcore/health" >/dev/null 2>&1 \
    || curl -fsS "http://127.0.0.1:${HOST_PORT}/" >/dev/null 2>&1) \
    && echo "✅ Local curl OK" \
    || echo "⚠️ Local curl FAILED (check logs below)"
else
  echo "⚠️ curl not found; skipping HTTP check"
fi

echo "🧾 Last container logs (tail)..."
sudo docker logs --tail 120 "${SAFE_IMAGE_NAME}" 2>&1 || true

# Reload Nginx nếu có
if command -v nginx >/dev/null 2>&1; then
  echo "🔄 Reloading Nginx..."
  sudo nginx -t && sudo systemctl reload nginx
fi

echo "✅ Deploy xong: ${SAFE_IMAGE_NAME} chạy trên cổng ${HOST_PORT}"