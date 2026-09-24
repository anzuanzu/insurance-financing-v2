FROM python:3.12-slim

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    SOFFICE_PATH=/usr/bin/soffice \
    CALCULATION_TIMEOUT_SECONDS=150 \
    INSURANCE_DATA_DIR=/data

RUN apt-get update \
    && apt-get install --no-install-recommends -y libreoffice-calc \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app
COPY requirements.txt ./
RUN pip install --no-cache-dir -r requirements.txt
COPY . ./

EXPOSE 8080
CMD ["sh", "-c", "python insurance_calculation_server.py --host 0.0.0.0 --port ${PORT:-8080}"]
