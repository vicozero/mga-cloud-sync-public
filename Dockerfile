FROM python:3.12-slim

WORKDIR /app

ENV PYTHONUNBUFFERED=1 \
    PYTHONDONTWRITEBYTECODE=1 \
    PIP_NO_CACHE_DIR=1

COPY requirements-cloud.txt .
RUN pip install --no-cache-dir -r requirements-cloud.txt

COPY . .

EXPOSE 8080

CMD ["sh", "-c", "uvicorn render_backend.main:app --host 0.0.0.0 --port ${PORT:-8080}"]