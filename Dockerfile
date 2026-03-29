# medaccur Playwright PPTX Renderer
# Optimised for Railway deployment

FROM python:3.12-slim

# System deps for Playwright Chromium
RUN apt-get update && apt-get install -y \
    wget curl gnupg ca-certificates \
    # Chromium runtime deps
    libnss3 libatk1.0-0 libatk-bridge2.0-0 \
    libcups2 libdrm2 libxkbcommon0 libxcomposite1 \
    libxdamage1 libxrandr2 libgbm1 libasound2 \
    libpango-1.0-0 libcairo2 libxshmfence1 \
    libx11-xcb1 libxcb-dri3-0 \
    fonts-liberation fonts-noto fonts-noto-cjk \
    # Cleanup
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# Install Python deps
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Install Playwright Chromium only (smallest footprint)
RUN playwright install chromium
RUN playwright install-deps chromium

# Copy app
COPY . .

# Railway uses PORT env var
ENV PORT=8080
EXPOSE 8080

CMD ["uvicorn", "main:app", "--host", "0.0.0.0", "--port", "8080", "--workers", "2"]
