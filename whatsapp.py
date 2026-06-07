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
import time
from pathlib import Path

import requests


XLSX_MIMETYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


# As variaveis sao lidas a cada chamada (e nao no import do modulo) de proposito:
# assim funciona mesmo que este modulo seja importado antes do load_dotenv().
def _url():
    return os.getenv("EVOLUTION_API_URL", "").rstrip("/")


def _key():
    return os.getenv("EVOLUTION_API_KEY", "")


def _instance():
    return os.getenv("EVOLUTION_INSTANCE", "")


def _destino():
    return os.getenv("WHATSAPP_DESTINO", "")


def configurado():
    """True se todas as credenciais necessarias estao no ambiente."""
    return all([_url(), _key(), _instance(), _destino()])


def _headers():
    return {"apikey": _key(), "Content-Type": "application/json"}


def enviar_texto(texto):
    url = f"{_url()}/message/sendText/{_instance()}"
    resp = requests.post(
        url,
        json={"number": _destino(), "text": texto},
        headers=_headers(),
        timeout=30,
    )
    resp.raise_for_status()


def enviar_arquivo(caminho, legenda=""):
    caminho = Path(caminho)
    conteudo_b64 = base64.b64encode(caminho.read_bytes()).decode("ascii")
    url = f"{_url()}/message/sendMedia/{_instance()}"
    payload = {
        "number": _destino(),
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
    except Exception as exc:
        return False, f"Falha ao enviar o resumo pelo WhatsApp: {exc}"

    # Envia cada anexo de forma independente: se um falhar, os demais ainda vao.
    # Uma pausa curta entre envios evita throttle do Baileys ao mandar varios.
    candidatos = [a for a in (arquivos or []) if Path(a).exists() and Path(a).stat().st_size > 0]
    enviados = 0
    falhas = []
    for arq in candidatos:
        try:
            enviar_arquivo(arq)
            enviados += 1
            time.sleep(1.5)
        except Exception as exc:
            falhas.append(f"{Path(arq).name} ({exc})")

    if not falhas:
        return True, f"Resumo + {enviados} planilha(s) enviados pelo WhatsApp."
    return False, (
        f"Resumo enviado; {enviados} de {len(candidatos)} planilha(s) foram. "
        f"Falharam: {'; '.join(falhas)}"
    )
