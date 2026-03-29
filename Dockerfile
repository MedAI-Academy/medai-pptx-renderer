# medaccur Playwright PPTX Renderer
# Official Playwright Python image — Chromium pre-installed, zero setup issues

FROM mcr.microsoft.com/playwright/python:v1.44.0-jammy

WORKDIR /app

# Python deps only — browser already in base image
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# App code
COPY . .

# Railway sets PORT dynamically; default 8080
ENV PORT=8080
EXPOSE 8080

# Shell form so $PORT is expanded at runtime
CMD uvicorn main:app --host 0.0.0.0 --port $PORT --workers 1
