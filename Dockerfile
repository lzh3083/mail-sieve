FROM python:3.11-slim

WORKDIR /app

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    TZ=Asia/Shanghai \
    MAIL_SIEVE_DB_PATH=/app/data/hotmail_inventory.sqlite3

COPY hotmail_manager_server.py /app/hotmail_manager_server.py
COPY hotmail-dashboard.html /app/hotmail-dashboard.html

RUN mkdir -p /app/data

EXPOSE 8899

HEALTHCHECK --interval=15s --timeout=5s --retries=5 \
  CMD python3 -c "import urllib.request; urllib.request.urlopen('http://127.0.0.1:8899/health', timeout=3)" || exit 1

CMD ["python3", "/app/hotmail_manager_server.py", "--host", "0.0.0.0", "--port", "8899"]
