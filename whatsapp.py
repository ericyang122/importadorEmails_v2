"""Envio de notificacoes pelo WhatsApp via Evolution API.

Plugavel: se as variaveis de ambiente nao estiverem configuradas, as funcoes
nao fazem nada (e nunca levantam excecao), entao a automacao roda normalmente
mesmo sem WhatsApp. Para ativar, preencha no .env:

    EVOLUTION_API_URL=https://sua-evolution.com   # URL base da sua Evolution API
    EVOLUTION_API_KEY=sua_api_key                 # apikey da instancia
    EVOLUTION_INSTANCE=nome_da_instancia          # nome da instancia conectada
    WHATSAPP_DESTINO=5511999998888                # numero que recebe o resumo (com DDI)

Para mandar pra mais de uma pessoa, separe os numeros por virgula:

    WHATSAPP_DESTINO=5511999998888,5511888887777

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


def _destinos():
    bruto = os.getenv("WHATSAPP_DESTINO", "")
    return [n.strip() for n in bruto.split(",") if n.strip()]


def _eh_grupo(destino):
    """Grupos do WhatsApp terminam em @g.us; numeros individuais nao."""
    return destino.strip().endswith("@g.us")


def configurado():
    """True se todas as credenciais necessarias estao no ambiente."""
    return all([_url(), _key(), _instance(), _destinos()])


def _headers():
    return {"apikey": _key(), "Content-Type": "application/json"}


def enviar_texto(texto, numero):
    url = f"{_url()}/message/sendText/{_instance()}"
    resp = requests.post(
        url,
        json={"number": numero, "text": texto},
        headers=_headers(),
        timeout=30,
    )
    resp.raise_for_status()


def enviar_arquivo(caminho, numero, legenda=""):
    caminho = Path(caminho)
    conteudo_b64 = base64.b64encode(caminho.read_bytes()).decode("ascii")
    url = f"{_url()}/message/sendMedia/{_instance()}"
    payload = {
        "number": numero,
        "mediatype": "document",
        "fileName": caminho.name,
        "media": conteudo_b64,
        "mimetype": XLSX_MIMETYPE,
    }
    if legenda:
        payload["caption"] = legenda
    resp = requests.post(url, json=payload, headers=_headers(), timeout=120)
    resp.raise_for_status()


def notificar(texto, arquivos=None, texto_grupo=None):
    """Envia o resumo e anexa os arquivos. Retorna (ok, mensagem).

    Silencioso e seguro: se nao houver configuracao, apenas informa. Qualquer
    erro de rede e capturado para nao derrubar a automacao.

    Se `texto_grupo` for informado, os destinos que forem GRUPO (@g.us) recebem
    essa versao (com saudacao, pro Marketing) e os numeros recebem `texto`.
    """
    if not configurado():
        return False, "WhatsApp nao configurado (.env) — notificacao ignorada."

    candidatos = [a for a in (arquivos or []) if Path(a).exists() and Path(a).stat().st_size > 0]
    destinos = _destinos()
    destinos_ok = 0
    falhas = []

    # Cada destinatario e cada anexo vao de forma independente: se um falhar,
    # os demais ainda vao. A pausa curta entre envios evita throttle do Baileys.
    for numero in destinos:
        msg = texto_grupo if (texto_grupo and _eh_grupo(numero)) else texto
        try:
            enviar_texto(msg, numero)
        except Exception as exc:
            falhas.append(f"resumo para {numero} ({exc})")
            continue

        destinos_ok += 1
        for arq in candidatos:
            try:
                enviar_arquivo(arq, numero)
                time.sleep(1.5)
            except Exception as exc:
                falhas.append(f"{Path(arq).name} para {numero} ({exc})")

    if not falhas:
        plural = f" para {destinos_ok} numero(s)" if len(destinos) > 1 else ""
        return True, f"Resumo + {len(candidatos)} planilha(s) enviados pelo WhatsApp{plural}."
    return False, (
        f"WhatsApp: resumo chegou em {destinos_ok} de {len(destinos)} numero(s). "
        f"Falharam: {'; '.join(falhas)}"
    )
