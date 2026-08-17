## Deploy lên VPS (Streamlit + Docker + GHCR)

Mục tiêu: chạy web public qua `http://160.191.50.186:6520`.

### 1) Chuẩn bị VPS
- **Mở firewall port 6520** (nếu VPS có bật UFW):

```bash
sudo ufw allow 6520/tcp
sudo ufw status
```

- **Cài Docker** (nếu chưa có) và đảm bảo chạy được `docker` + `docker compose`.

### 2) Secrets cần set trên GitHub
Vào GitHub repo → **Settings → Secrets and variables → Actions → New repository secret**.

#### A. Secrets cho deploy Docker/GHCR/VPS
- **`IMAGE_NAME`**: tên app/image (ví dụ `mailmerge` hoặc `cdimex-mailmerge`).
  - **Cách lấy value**: bạn tự chọn; nên là tên ngắn, không dấu, không khoảng trắng.

- **`VPS_HOST`**: IP VPS.
  - **Value**: `160.191.50.186`

- **`VPS_USER`**: user SSH trên VPS (ví dụ `root`, `ubuntu`, `deploy`).
  - **Cách lấy value**: user bạn dùng để SSH vào VPS (cùng user có quyền chạy docker).

- **`VPS_SSH_KEY`**: private key SSH dùng cho GitHub Actions SSH vào VPS.
  - **Cách lấy value** (khuyến nghị tạo key riêng):

```bash
ssh-keygen -t ed25519 -C "github-actions-deploy" -f ./vps_deploy_key -N ""
```

  - Copy **public key** `vps_deploy_key.pub` lên VPS vào `~/.ssh/authorized_keys` của `VPS_USER`.
  - Copy **private key** (nội dung file `vps_deploy_key`) dán nguyên văn vào secret `VPS_SSH_KEY`.
  - Workflow hỗ trợ key **raw hoặc base64**.

- **`GHCR_USERNAME`**: username GitHub sở hữu token (thường là username của bạn / account tạo token).
  - **Cách lấy value**: GitHub username của account tạo `GHCR_TOKEN`.

- **`GHCR_TOKEN`**: token để VPS `docker login ghcr.io` và pull image.
  - **Cách lấy value**:
    - GitHub → **Settings → Developer settings → Personal access tokens**.
    - Tạo token và cấp quyền tối thiểu:
      - Nếu dùng **classic PAT**: `read:packages` (và thêm `repo` nếu repo/private packages yêu cầu).
      - Nếu dùng **fine-grained token**: cấp quyền **read packages** cho đúng owner/repo.
    - Dán token vào secret `GHCR_TOKEN`.

#### B. Secrets cho port và env runtime
- **`HOST_PORT`**: port publish trên VPS.
  - **Value**: `6520`

- **`CONTAINER_PORT`**: port app bên trong container.
  - **Value**: `6520`

- **`PROD_ENV_FILE`**: nội dung file `.env` sẽ được ghi lên VPS và nạp vào container.
  - **Cách lấy value**:
    - Copy nội dung từ `.env.example`, thay các biến (đặc biệt `SMTP_USER`, `SMTP_PASS`) rồi paste nguyên khối vào secret này.

#### C. Optional (hiện workflow có khai báo)
- **`DOMAIN`**: nếu bạn dùng Nginx/domain. Hiện deploy script không bắt buộc.
  - **Cách lấy value**: domain của bạn (ví dụ `mail.example.com`) hoặc đặt tạm `-`.

### 3) Deploy
- Push code lên branch `main`/`master` → GitHub Actions sẽ tự build & deploy.
- Truy cập: `http://160.191.50.186:6520`

