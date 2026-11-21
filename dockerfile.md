FROM python:3.11-slim-bookworm

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PLAYWRIGHT_BROWSERS_PATH=/ms-playwright

WORKDIR /app

# Install Git, LibreOffice Calc (for headless Excel formula recalculation), and system utilities
RUN apt-get update && apt-get install -y --no-install-recommends \
    git \
    libreoffice-calc \
    libreoffice-core \
    ca-certificates \
    curl \
    && rm -rf /var/lib/apt/lists/*

# Install Python dependencies first for better Docker layer caching
COPY requirements.txt /app/requirements.txt
RUN pip install --no-cache-dir -r /app/requirements.txt

# Install Playwright Chromium browser + OS dependencies for EvanMiya / JS scrapers
RUN playwright install --with-deps chromium

# Copy repository files (including scripts, Excel workbooks, and dashboard HTML)
COPY . /app

EXPOSE 10000

CMD ["python3", "cbb_local_server.py"]
