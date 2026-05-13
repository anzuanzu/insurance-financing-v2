FROM python:3.11-slim

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1

WORKDIR /app

RUN apt-get update \
    && apt-get install -y --no-install-recommends libreoffice-calc \
    && rm -rf /var/lib/apt/lists/*

COPY backend/requirements.txt /app/requirements.txt
RUN pip install --no-cache-dir -r /app/requirements.txt

COPY backend /app/backend
COPY scripts /app/scripts
COPY "WUS(金控)國泰人壽鑽美滿利率變動型美元終身壽險(定期給付型)-11501版(宣告4.25%-1150101啟用)(保費融資)_i值.xlsx" /app/

ENV PYTHONPATH=/app/backend:/app

EXPOSE 8000

CMD ["uvicorn", "app.main:app", "--host", "0.0.0.0", "--port", "8000"]
