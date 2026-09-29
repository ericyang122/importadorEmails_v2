# -*- coding: utf-8 -*-
"""Texto e ordem dos anexos do aviso de WhatsApp no fim de cada automacao.

Saiu do app.py (30/09) porque agora dois lugares mandam esse aviso: o app na VM e o
executor do Cloud Run Job (executor_job.py). Mesma mensagem nos dois.
"""
from datetime import datetime
from pathlib import Path


def _formatar_duracao(started_at, finished_at):
    try:
        segundos = int((datetime.fromisoformat(finished_at) - datetime.fromisoformat(started_at)).total_seconds())
    except (TypeError, ValueError):
        return "tempo desconhecido"
    segundos = max(0, segundos)
    horas, resto = divmod(segundos, 3600)
    minutos, seg = divmod(resto, 60)
    if horas:
        return f"{horas}h{minutos:02d}min"
    if minutos:
        return f"{minutos}min{seg:02d}s"
    return f"{seg}s"


def montar_resumo(job):
    """Resumo curto do WhatsApp (28/09: o grupo pediu menos mensagem).

    Vai como legenda da primeira planilha; o detalhe esta nas planilhas.
    """
    progress = job.get("progress") or {}
    modo = job.get("mode")
    rotulo_sucesso, rotulo_pendente = {
        "consulta": ("encontrados", "não encontrados"),
        "verificar": ("com cadastro", "sem cadastro"),
    }.get(modo, ("cadastrados", "não cadastrados"))

    total = progress.get("total", 0) or 0
    processados = progress.get("processados", 0) or 0
    sucessos = progress.get("sucessos", 0) or 0
    pendentes = progress.get("pendentes", 0) or 0
    erros = progress.get("erros", 0) or 0
    nome = job.get("filename", "")

    status = job.get("status")
    if status == "completed":
        titulo = f"✅ *{nome}* pronta"
    elif status == "stopped":
        titulo = f"⏸️ *{nome}* parada em {processados}/{total}"
    else:
        titulo = f"❌ *{nome}* deu erro em {processados}/{total}"
    numeros = f"{sucessos} {rotulo_sucesso} · {pendentes} {rotulo_pendente}"
    if erros:
        numeros += f" · {erros} com erro de consulta"
    return f"{titulo}\n{numeros}"


def ordenar_anexos(job, arquivos):
    """Encontrados primeiro, depois nao encontrados, erros por ultimo.

    A planilha de erros so vai se tiver erro (sem erro ela vai so com cabecalho).
    """
    erros = ((job.get("progress") or {}).get("erros", 0) or 0)

    def peso(caminho):
        nome = Path(caminho).name.lower()
        if "erro" in nome:
            return 2
        if "sem_" in nome or "nao_" in nome or "não_" in nome:
            return 1
        return 0

    anexos = [a for a in arquivos if Path(a).exists() and Path(a).stat().st_size > 0]
    if not erros:
        anexos = [a for a in anexos if peso(a) != 2]
    return sorted(anexos, key=peso)
