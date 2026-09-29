# -*- coding: utf-8 -*-
"""Executor do importador no Cloud Run Job — faz o que o run_automation() do app.py
faz na VM, mas lendo e gravando no bucket (ver nuvem.py pro desenho das pastas).

Recebe JOB_ID por variável de ambiente (o app dispara a execução com esse override).
Roda o confio_api.py como subprocesso — o MESMO robô da VM, sem mudar uma linha —,
repassa progresso e log pro estado.json a cada poucos segundos, obedece o "parar"
da tela, sobe os resultados e manda o aviso no WhatsApp no fim.
"""
import json
import os
import subprocess
import sys
import tempfile
import time
from pathlib import Path

from dotenv import load_dotenv

BASE_DIR = Path(__file__).resolve().parent
load_dotenv(os.getenv("IMPORTADOR_ENV_FILE") or BASE_DIR / ".env")

import nuvem  # noqa: E402
import resumo  # noqa: E402
import whatsapp  # noqa: E402

MAX_LOG_LINES = 2000
INTERVALO_ESTADO_S = 5
INTERVALO_PARAR_S = 10


def sanitiza(linha, segredos):
    for s in segredos:
        if s:
            linha = linha.replace(s, "[oculto]")
    return linha


def main():
    job_id = os.environ["JOB_ID"]
    meta = nuvem.le_json(nuvem.caminho_job(job_id, "meta.json"))
    estado = nuvem.le_json(nuvem.caminho_job(job_id, "estado.json"), {})
    if not meta:
        print(f"job {job_id} sem meta.json — nada a fazer")
        return 1
    cred = nuvem.le_json(nuvem.caminho_job(job_id, "credenciais.json"), {})
    nuvem.apaga(nuvem.caminho_job(job_id, "credenciais.json"))   # senha não fica no bucket
    login, senha = cred.get("login", ""), cred.get("senha", "")
    segredos = [login, senha]
    slug = meta["slug"]

    def grava_estado(**extra):
        estado.update(extra)
        estado["batimento"] = nuvem.agora_iso()
        estado["logs"] = estado.get("logs", [])[-MAX_LOG_LINES:]
        nuvem.grava_json(nuvem.caminho_job(job_id, "estado.json"), estado)

    def loga(linha):
        estado.setdefault("logs", []).append(sanitiza(linha, segredos))

    if not login or not senha:
        loga("\nErro ao executar automacao: credenciais do Sigavi nao chegaram ao executor.\n")
        grava_estado(status="failed", finished_at=nuvem.agora_iso())
        return 1

    tmp = Path(tempfile.mkdtemp(prefix=f"imp_{job_id[:8]}_"))
    excel = tmp / meta["secure_name"]
    nuvem.baixa_arquivo(nuvem.caminho_backup(slug, f"entrada_{meta['secure_name']}"), excel)
    result_dir = tmp / "resultados"
    result_dir.mkdir()
    progresso = tmp / "progresso.json"
    parar_local = tmp / "parar-e-salvar.flag"

    comando = [sys.executable, "-u", str(BASE_DIR / "confio_api.py"), "--excel", str(excel),
               "--result-dir", str(result_dir), "--progress-file", str(progresso),
               "--stop-file", str(parar_local), "--mode", meta["mode"]]
    env = os.environ.copy()
    env.update({"SIGAVI_LOGIN": login, "SIGAVI_SENHA": senha, "PYTHONIOENCODING": "utf-8"})
    env.pop("SIGAVI_EMPREENDIMENTO_TELA", None)
    if meta.get("empreendimento"):
        env["SIGAVI_EMPREENDIMENTO_TELA"] = meta["empreendimento"]

    loga("Automacao iniciada.\n")
    grava_estado(status="running", started_at=nuvem.agora_iso())
    rc = None
    try:
        proc = subprocess.Popen(comando, cwd=BASE_DIR, env=env, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                                text=True, encoding="utf-8", errors="replace", bufsize=1)
        ultimo_estado = ultimo_parar = time.time()
        for linha in proc.stdout:
            if linha.startswith("PROGRESS="):
                try:
                    estado["progress"] = json.loads(linha[len("PROGRESS="):])
                except ValueError:
                    pass
            else:
                loga(linha)
            agora = time.time()
            if agora - ultimo_parar > INTERVALO_PARAR_S:
                ultimo_parar = agora
                if not parar_local.exists() and nuvem.existe(nuvem.caminho_job(job_id, "parar")):
                    parar_local.write_text("parar")
                    loga("\nParada solicitada. Salvando resultados no proximo ponto seguro...\n")
                    estado["status"] = "stopping"
            if agora - ultimo_estado > INTERVALO_ESTADO_S:
                ultimo_estado = agora
                grava_estado()
        rc = proc.wait()
        if rc == 0 and parar_local.exists():
            loga("\nAutomacao parada com resultados salvos.\n")
            status = "stopped"
        elif rc == 0:
            loga("\nAutomacao concluida.\n")
            status = "completed"
        else:
            loga(f"\nAutomacao finalizada com erro. Codigo: {rc}\n")
            status = "failed"
    except Exception as exc:
        loga(f"\nErro ao executar automacao: {exc}\n")
        status = "failed"

    # resultados + progresso + log no backup (o reprocessar lê daqui)
    arquivos = sorted(result_dir.glob("*.xlsx"), key=lambda p: p.name)
    entradas = []
    for i, arq in enumerate(arquivos):
        destino = nuvem.caminho_backup(slug, f"resultados/{arq.name}")
        nuvem.sobe_arquivo(arq, destino)
        entradas.append({"id": str(i), "filename": arq.name, "gcs": destino, "downloaded": False})
    if progresso.exists():
        nuvem.sobe_arquivo(progresso, nuvem.caminho_backup(slug, "progresso.json"))
    grava_estado(status=status, return_code=rc, finished_at=nuvem.agora_iso(),
                 result_files=entradas, download_available=bool(entradas))

    # aviso no WhatsApp (mesmo texto da VM)
    if whatsapp.credenciais_ok():
        anexos = resumo.ordenar_anexos(estado, [str(a) for a in arquivos])
        ok, msg = whatsapp.notificar(resumo.montar_resumo(estado), anexos,
                                     destinos=meta.get("destinos") or None, legenda_no_anexo=True)
        loga(f"\n{msg}\n")
    nuvem.grava_bytes(nuvem.caminho_backup(slug, "log.txt"), "".join(estado.get("logs", [])).encode("utf-8"),
                      "text/plain; charset=utf-8")
    grava_estado()
    ativo = nuvem.le_json(f"{nuvem.PREFIXO}/ativo.json", {})
    if ativo.get("job_id") == job_id:
        nuvem.apaga(f"{nuvem.PREFIXO}/ativo.json")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
