# Importador Sigavi no Cloud Run (30/09/2026). A MESMA imagem serve os dois papéis:
#   serviço (tela + API):  waitress no $PORT                     — padrão (CMD abaixo)
#   Job (a automação):     python executor_job.py  (o Job sobrescreve o comando)
# Sem navegador: o motor é o confio_api.py (API REST do Sigavi). curl é obrigatório —
# o sigavi_api.py fala com o IIS do Sigavi por curl (urllib/fetch tomam 500).
FROM python:3.12-slim
ENV PYTHONUNBUFFERED=1 PYTHONIOENCODING=utf-8 TZ=America/Sao_Paulo \
    EXECUCAO=cloudrun IMPORTADOR_MOTOR=api API_SOMENTE_LOCAL=false SESSION_COOKIE_SECURE=true
RUN apt-get update && apt-get install -y --no-install-recommends curl tzdata ca-certificates \
    && rm -rf /var/lib/apt/lists/*
WORKDIR /app
COPY requirements.txt requirements-cloud.txt ./
RUN pip install --no-cache-dir -r requirements-cloud.txt
COPY . .
CMD exec waitress-serve --listen=0.0.0.0:${PORT:-8080} --threads=8 app:app
