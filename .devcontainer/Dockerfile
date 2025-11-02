# ---------- 1. Build-Stage ----------
FROM python:3.11-slim AS builder
WORKDIR /app
COPY requirements.txt .
RUN pip install --upgrade pip && \
    pip install --user -r requirements.txt

# ---------- 2. Runtime-Stage ----------
FROM python:3.11-slim
WORKDIR /app
ENV PYTHONPATH=/app

# 1) Python-Pakete aus Build-Stage übernehmen
COPY --from=builder /root/.local /root/.local
ENV PATH=/root/.local/bin:$PATH

# 2) Git für VS-Code nachinstallieren  << NEU >>
RUN apt-get update && \
    apt-get install -y --no-install-recommends git && \
    rm -rf /var/lib/apt/lists/*

# ---------- Projektdateien ----------
COPY Bartek/output /app/Bartek/output
COPY Bartek/input  /app/Bartek/input
COPY Arno/output   /app/Arno/output
COPY Arno/input    /app/Arno/input

USER root

# --------- Default Entry Point -------
ENTRYPOINT ["python", "Bartek/output/run_calc.py"]
    