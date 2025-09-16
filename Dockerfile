FROM python:3.11-slim

# System deps (build + runtime). Add poppler-utils if you process PDFs.
RUN apt-get update -y \
    && apt-get install -y --no-install-recommends \
       build-essential \
       curl \
       poppler-utils \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# Install Python deps first for better caching
COPY requirements.txt /app/
RUN pip install --no-cache-dir -r requirements.txt

# Copy source code
COPY . /app

# Streamlit specific: configure to listen on 0.0.0.0:7001
ENV STREAMLIT_SERVER_PORT=7001 \
    STREAMLIT_SERVER_HEADLESS=true \
    STREAMLIT_BROWSER_GATHER_USAGE_STATS=false

EXPOSE 7001

# Healthcheck: try to hit the root after container starts
HEALTHCHECK --interval=30s --timeout=5s --retries=5 CMD curl -fsS http://127.0.0.1:7001/ || exit 1

# Default command: run Streamlit app
CMD ["bash", "-lc", "streamlit run streamlit_app.py --server.address=0.0.0.0 --server.port=${STREAMLIT_SERVER_PORT}"]


