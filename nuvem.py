# -*- coding: utf-8 -*-
"""Importador no Cloud Run (30/09/2026): tela = serviço, automação = Cloud Run Job.

Por que Job: a automação leva minutos e o Cloud Run desliga serviço sem requisição —
dentro do app (como na VM) ela morreria no meio quando ninguém estivesse olhando a tela
(ex.: pedido do eriquizinho). O Job roda até o fim.

Tudo que era pasta local vira objeto no bucket, com o MESMO desenho da VM:
    importador/jobs/<job_id>/meta.json          o que a tela decidiu (modo, destinos, empreendimento...)
    importador/jobs/<job_id>/credenciais.json   login/senha do Sigavi digitados — o Job APAGA ao ler
    importador/jobs/<job_id>/estado.json        status + progresso + log (a tela lê daqui)
    importador/jobs/<job_id>/parar              existe = pediram pra parar
    importador/backups/<slug>/entrada_X.xlsx, progresso.json, log.txt, resultados/*.xlsx
    importador/ativo.json                       qual job está rodando (trava de 1 por vez)

Só é usado com EXECUCAO=cloudrun. Na VM o app.py segue com a thread local de sempre.
"""
import json
import os
import time
from datetime import datetime, timezone

import google.auth
import google.auth.transport.requests
import requests
from google.cloud import storage

BUCKET = os.getenv("IMPORTADOR_BUCKET", "laik-staging-abyara-automacoes")
PROJETO = os.getenv("GOOGLE_CLOUD_PROJECT", "laik-staging")
REGIAO = os.getenv("IMPORTADOR_REGIAO", "southamerica-east1")
JOB_NOME = os.getenv("IMPORTADOR_JOB", "abyara-importador-exec")
PREFIXO = "importador"
# job sem notícia há mais que isso é considerado morto (o Job atualiza a cada ~5s)
BATIMENTO_MAX_S = int(os.getenv("IMPORTADOR_BATIMENTO_MAX_S", "600"))

_cliente = None


def _bucket():
    global _cliente
    if _cliente is None:
        _cliente = storage.Client(project=PROJETO)
    return _cliente.bucket(BUCKET)


def agora_iso():
    return datetime.now().isoformat(timespec="seconds")


# ── objetos ───────────────────────────────────────────────────────────────────
def caminho_job(job_id, nome):
    return f"{PREFIXO}/jobs/{job_id}/{nome}"


def caminho_backup(slug, nome=""):
    return f"{PREFIXO}/backups/{slug}/{nome}".rstrip("/")


def grava_json(caminho, dados):
    _bucket().blob(caminho).upload_from_string(
        json.dumps(dados, ensure_ascii=False), content_type="application/json")


def le_json(caminho, padrao=None):
    blob = _bucket().blob(caminho)
    try:
        return json.loads(blob.download_as_bytes())
    except Exception:
        return padrao


def grava_bytes(caminho, dados, tipo="application/octet-stream"):
    _bucket().blob(caminho).upload_from_string(dados, content_type=tipo)


def le_bytes(caminho):
    return _bucket().blob(caminho).download_as_bytes()


def sobe_arquivo(caminho_local, caminho):
    _bucket().blob(caminho).upload_from_filename(str(caminho_local))


def baixa_arquivo(caminho, caminho_local):
    _bucket().blob(caminho).download_to_filename(str(caminho_local))


def existe(caminho):
    return _bucket().blob(caminho).exists()


def apaga(caminho):
    try:
        _bucket().blob(caminho).delete()
    except Exception:
        pass


def lista(prefixo):
    _bucket()
    return [b.name for b in _cliente.list_blobs(BUCKET, prefix=prefixo)]


# ── trava de 1 automação por vez ──────────────────────────────────────────────
def job_ativo():
    """job_id do que está rodando (com batimento recente), ou None."""
    ativo = le_json(f"{PREFIXO}/ativo.json")
    if not ativo or not ativo.get("job_id"):
        return None
    estado = le_json(caminho_job(ativo["job_id"], "estado.json"), {})
    if estado.get("status") not in ("queued", "running", "stopping"):
        return None
    try:
        batida = datetime.fromisoformat(estado.get("batimento") or estado.get("created_at"))
    except (TypeError, ValueError):
        return None
    if (datetime.now() - batida).total_seconds() > BATIMENTO_MAX_S:
        return None          # o Job morreu sem avisar: não trava a fila pra sempre
    return ativo["job_id"]


def marca_ativo(job_id):
    grava_json(f"{PREFIXO}/ativo.json", {"job_id": job_id, "desde": agora_iso()})


# ── disparo do Cloud Run Job ──────────────────────────────────────────────────
def _token():
    cred, _ = google.auth.default(scopes=["https://www.googleapis.com/auth/cloud-platform"])
    cred.refresh(google.auth.transport.requests.Request())
    return cred.token


def dispara_execucao(job_id):
    """Roda o Job com JOB_ID do pedido. Devolve o nome da execução."""
    url = f"https://run.googleapis.com/v2/projects/{PROJETO}/locations/{REGIAO}/jobs/{JOB_NOME}:run"
    corpo = {"overrides": {"containerOverrides": [{"env": [{"name": "JOB_ID", "value": job_id}]}]}}
    r = requests.post(url, json=corpo, headers={"Authorization": f"Bearer {_token()}"}, timeout=30)
    if r.status_code >= 300:
        raise RuntimeError(f"nao consegui disparar o Job ({r.status_code}): {r.text[:200]}")
    return (r.json().get("metadata") or {}).get("name", "")
