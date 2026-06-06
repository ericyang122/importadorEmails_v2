"""Envio de notificacoes pelo WhatsApp via Evolution API.

Plugavel: se as variaveis de ambiente nao estiverem configuradas, as funcoes
nao fazem nada (e nunca levantam excecao), entao a automacao roda normalmente
mesmo sem WhatsApp. Para ativar, preencha no .env:

    EVOLUTION_API_URL=https://sua-evolution.com   # URL base da sua Evolution API
    EVOLUTION_API_KEY=sua_api_key                 # apikey da instancia
    EVOLUTION_INSTANCE=nome_da_instancia          # nome da instancia conectada
    WHATSAPP_DESTINO=5511999998888                # numero que recebe o resumo (com DDI)

Compativel com Evolution API v2 (endpoints /message/sendText e /message/sendMedia).
"""

import base64
import os
from pathlib import Path

import requests


EVOLUTION_API_URL = os.getenv("EVOLUTION_API_URL", "").rstrip("/")
EVOLUTION_API_KEY = os.getenv("EVOLUTION_API_KEY", "")
EVOLUTION_INSTANCE = os.getenv("EVOLUTION_INSTANCE", "")
WHATSAPP_DESTINO = os.getenv("WHATSAPP_DESTINO", "")

XLSX_MIMETYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


def configurado():
    """True se todas as credenciais necessarias estao no ambiente."""
    return all([EVOLUTION_API_URL, EVOLUTION_API_KEY, EVOLUTION_INSTANCE, WHATSAPP_DESTINO])


def _headers():
    return {"apikey": EVOLUTION_API_KEY, "Content-Type": "application/json"}


def enviar_texto(texto):
    url = f"{EVOLUTION_API_URL}/message/sendText/{EVOLUTION_INSTANCE}"
    resp = requests.post(
        url,
        json={"number": WHATSAPP_DESTINO, "text": texto},
        headers=_headers(),
        timeout=30,
    )
    resp.raise_for_status()


def enviar_arquivo(caminho, legenda=""):
    caminho = Path(caminho)
    conteudo_b64 = base64.b64encode(caminho.read_bytes()).decode("ascii")
    url = f"{EVOLUTION_API_URL}/message/sendMedia/{EVOLUTION_INSTANCE}"
    payload = {
        "number": WHATSAPP_DESTINO,
        "mediatype": "document",
        "fileName": caminho.name,
        "media": conteudo_b64,
        "mimetype": XLSX_MIMETYPE,
    }
    if legenda:
        payload["caption"] = legenda
    resp = requests.post(url, json=payload, headers=_headers(), timeout=120)
    resp.raise_for_status()


def notificar(texto, arquivos=None):
    """Envia o resumo e anexa os arquivos. Retorna (ok, mensagem).

    Silencioso e seguro: se nao houver configuracao, apenas informa. Qualquer
    erro de rede e capturado para nao derrubar a automacao.
    """
    if not configurado():
        return False, "WhatsApp nao configurado (.env) — notificacao ignorada."
    try:
        enviar_texto(texto)
        for arq in (arquivos or []):
            if Path(arq).exists() and Path(arq).stat().st_size > 0:
                enviar_arquivo(arq)
        return True, "Resumo enviado pelo WhatsApp."
    except Exception as exc:
        return False, f"Falha ao enviar WhatsApp: {exc}"
