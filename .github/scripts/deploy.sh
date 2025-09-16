#!/usr/bin/env bash
set -euo pipefail

SAFE_IMAGE_NAME=$(echo "${IMAGE_NAME:-app}" | tr '[:upper:]' '[:lower:]' | sed -E 's/[^a-z0-9._-]+/-/g')
APP_DIR="/opt/${SAFE_IMAGE_NAME}"
HOST_PORT="${HOST_PORT:-9000}"
CONTAINER_PORT="${CONTAINER_PORT:-7001}"

sudo mkdir -p "$APP_DIR"

# Write .env
if [ -n "${PROD_ENV_FILE:-}" ]; then
  printf "%s" "${PROD_ENV_FILE}" | sudo tee "$APP_DIR/.env" > /dev/null
else
  sudo touch "$APP_DIR/.env"
fi

# Ensure docker compose
if ! docker compose version >/dev/null 2>&1; then
  if [ -f /etc/os-release ]; then
    . /etc/os-release
    if [ "${ID}" = "ubuntu" ] || [ "${ID}" = "debian" ]; then
      sudo apt-get update -y
      sudo apt-get install -y docker-compose-plugin
    fi
  fi
fi

# Login GHCR
echo "${GHCR_TOKEN}" | sudo docker login ghcr.io -u "${GHCR_USERNAME}" --password-stdin

# docker-compose.yml
sudo tee "$APP_DIR/docker-compose.yml" > /dev/null <<YAML
services:
  app:
    image: ${REPO_PATH}:${TAG_SHA}
    container_name: ${SAFE_IMAGE_NAME}
    restart: unless-stopped
    env_file:
      - .env
    ports:
      - "127.0.0.1:${HOST_PORT}:${CONTAINER_PORT}"
YAML

export COMPOSE_PROJECT_NAME="${SAFE_IMAGE_NAME}"
sudo -E docker compose -f "$APP_DIR/docker-compose.yml" pull app
sudo -E docker compose -f "$APP_DIR/docker-compose.yml" up -d --remove-orphans