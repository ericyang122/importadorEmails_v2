import io
import json
import os
import re
import secrets
import shutil
import subprocess
import sys
import threading
import uuid
from datetime import datetime
from decimal import Decimal, InvalidOperation
from functools import wraps
from pathlib import Path

from dotenv import load_dotenv
from flask import Flask, Response, jsonify, redirect, render_template, request, session, url_for
from werkzeug.utils import secure_filename

import whatsapp


BASE_DIR = Path(__file__).resolve().parent
UPLOAD_ROOT = BASE_DIR / "uploads"
BACKUP_ROOT = BASE_DIR / "backups"
MAX_LOG_LINES = 2000
RUNNING_STATUSES = {"queued", "running", "stopping"}
XLSX_MIMETYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

load_dotenv(BASE_DIR / ".env")
UPLOAD_ROOT.mkdir(exist_ok=True)
BACKUP_ROOT.mkdir(exist_ok=True)

DESTINOS_FILE = BASE_DIR / "destinos.json"


def carregar_destinos():
    """Lista curada de destinos do WhatsApp (nome + id), lida de destinos.json.

    Se o arquivo nao existir/for invalido, cai nos numeros do WHATSAPP_DESTINO
    (.env), sem nome amigavel. Cada item: {nome, id, padrao}.
    """
    try:
        dados = json.loads(DESTINOS_FILE.read_text(encoding="utf-8"))
        destinos = []
        for item in dados:
            wid = str(item.get("id", "")).strip()
            if wid:
                destinos.append({
                    "nome": str(item.get("nome", "")).strip() or wid,
                    "id": wid,
                    "padrao": bool(item.get("padrao", False)),
                })
        return destinos
    except (OSError, ValueError):
        brutos = [n.strip() for n in os.getenv("WHATSAPP_DESTINO", "").split(",") if n.strip()]
        return [{"nome": n, "id": n, "padrao": True} for n in brutos]


def validar_destinos(ids):
    """Mantem so os ids que existem na lista curada (allowlist), sem duplicar."""
    permitidos = {d["id"] for d in carregar_destinos()}
    return [i for i in dict.fromkeys(ids) if i in permitidos]

app = Flask(__name__)
app.secret_key = os.getenv("APP_SECRET_KEY") or secrets.token_hex(32)
app.config["MAX_CONTENT_LENGTH"] = int(os.getenv("MAX_UPLOAD_MB", "25")) * 1024 * 1024
app.config.update(
    SESSION_COOKIE_HTTPONLY=True,
    SESSION_COOKIE_SAMESITE="Lax",
    SESSION_COOKIE_SECURE=os.getenv("SESSION_COOKIE_SECURE", "").strip().lower()
    in {"1", "true", "yes", "sim"},
)
# Recarrega templates e nao deixa o navegador cachear CSS/JS durante o uso:
# evita o descompasso "HTML velho + CSS novo" ao atualizar a interface.
app.config["TEMPLATES_AUTO_RELOAD"] = True
app.config["SEND_FILE_MAX_AGE_DEFAULT"] = 0

# Mesma normalizacao de colunas usada no confio.py, para a previa bater com a execucao.
COLUNAS_RENOMEAR = {
    "CORRETOR ORIGEM": "CORRETOR DE ORIGEM",
    "TELEFONE": "FONE2",
    "NOME COMPLETO": "NOME",
    "nome_cliente": "NOME",
    "celular": "FONE2",
    "corretor": "CORRETOR DE ORIGEM",
    "origem": "TIPO PLANTAO",
    "gerente": "GERENTE",
}

# Criterios de busca do modo Consulta (mesma deteccao do confio.py): a planilha
# pode trazer o numero da FAC, o email ou so o nome do cliente.
RE_FAC_COL = re.compile(r"^\s*(fac|n[º°o]?\.?\s*fac|numero|n[º°o])\s*$", re.IGNORECASE)
RE_NOME_COL = re.compile(r"(nome|cliente)", re.IGNORECASE)
RE_EMAIL_COL = re.compile(r"e.?mail", re.IGNORECASE)


def _detectar_coluna(df, regex, full=False):
    """Primeira coluna cujo nome casa com o regex (fullmatch ou search)."""
    for c in df.columns:
        s = str(c)
        if (regex.fullmatch(s) if full else regex.search(s)):
            return c
    return None

APP_PASSWORD = os.getenv("APP_PASSWORD")
if not APP_PASSWORD:
    raise RuntimeError("APP_PASSWORD nao configurado — recusando iniciar (sem senha padrao insegura).")
JOBS = {}
JOBS_LOCK = threading.Lock()
RUN_LOCK = threading.Lock()


def get_csrf_token():
    if "csrf_token" not in session:
        session["csrf_token"] = secrets.token_urlsafe(32)
    return session["csrf_token"]


@app.context_processor
def inject_template_values():
    return {"csrf_token": get_csrf_token()}


def check_csrf():
    return request.form.get("csrf_token") == session.get("csrf_token")


def normalizar_telefone(valor):
    if valor is None:
        return ""
    texto = str(valor).strip()
    if texto.lower() in {"", "nan", "none"}:
        return ""
    try:
        numero = Decimal(texto)
        if numero.is_finite() and numero == numero.to_integral_value():
            return str(int(numero))
    except InvalidOperation:
        pass
    return re.sub(r"\D", "", texto)


def ler_planilha_upload(upload):
    import pandas as pd

    conteudo = upload.read()
    upload.stream.seek(0)
    excel = pd.ExcelFile(io.BytesIO(conteudo))
    sheet = excel.sheet_names[0]
    return excel.parse(sheet).rename(columns=COLUNAS_RENOMEAR), sheet


def dados_validos_planilha(df, mode):
    email_col = _detectar_coluna(df, RE_EMAIL_COL)
    if mode == "consulta":
        # vale a linha que tiver QUALQUER criterio de busca: FAC, email ou nome.
        fac_col = _detectar_coluna(df, RE_FAC_COL, full=True)
        nome_col = _detectar_coluna(df, RE_NOME_COL) or ("NOME" if "NOME" in df.columns else None)

        def _preenchido(col):
            if not col or col not in df.columns:
                return None
            return df[col].apply(lambda v: str(v).strip().lower() not in {"", "nan", "none"})

        mask = None
        for col in (fac_col, email_col, nome_col):
            m = _preenchido(col)
            if m is None:
                continue
            mask = m if mask is None else (mask | m)
        validos = int(mask.sum()) if mask is not None else 0
        return validos, email_col

    if "FONE2" not in df.columns:
        return 0, email_col
    validos = int(df["FONE2"].apply(lambda valor: len(normalizar_telefone(valor)) >= 11).sum())
    return validos, email_col


def login_required(route):
    @wraps(route)
    def wrapper(*args, **kwargs):
        if not session.get("authenticated"):
            return redirect(url_for("index"))
        return route(*args, **kwargs)

    return wrapper


def sanitize_log(line, secrets_to_hide):
    clean_line = line
    for secret in secrets_to_hide:
        if secret:
            clean_line = clean_line.replace(secret, "[oculto]")
    return clean_line


def backup_slug(filename):
    stem = secure_filename(Path(filename).stem) or "planilha"
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    return f"{timestamp}_{stem}_{uuid.uuid4().hex[:8]}"


def append_file_log(log_path, line, secrets_to_hide):
    clean_line = sanitize_log(line, secrets_to_hide)
    log_path.parent.mkdir(parents=True, exist_ok=True)
    with open(log_path, "a", encoding="utf-8") as log_file:
        log_file.write(clean_line)


def update_progress(job_id, line):
    """Le uma linha 'PROGRESS={...}' emitida pelo confio.py e guarda no job.

    Esses contadores sao a fonte de verdade do painel (substituem a contagem
    antiga feita lendo os logs, que perdia sucessos quando o buffer rotacionava).
    """
    try:
        data = json.loads(line[len("PROGRESS="):].strip())
    except (ValueError, TypeError):
        return
    with JOBS_LOCK:
        job = JOBS.get(job_id)
        if job:
            job["progress"] = data


def append_log(job_id, line, secrets_to_hide=None):
    secrets_to_hide = secrets_to_hide or []
    with JOBS_LOCK:
        job = JOBS.get(job_id)
        if not job:
            return
        job["logs"].append(sanitize_log(line, secrets_to_hide))
        if len(job["logs"]) > MAX_LOG_LINES:
            job["logs"] = job["logs"][-MAX_LOG_LINES:]


def has_active_job():
    with JOBS_LOCK:
        return any(job["status"] in RUNNING_STATUSES for job in JOBS.values())


def cleanup_finished_job_files():
    with JOBS_LOCK:
        finished_job_ids = [
            job_id
            for job_id, job in JOBS.items()
            if job["status"] not in RUNNING_STATUSES
        ]
        for job_id in finished_job_ids:
            job = JOBS[job_id]
            job["download_available"] = False
            job["result_files"] = []
            if not job.get("result_deleted_at"):
                job["result_deleted_at"] = datetime.now().isoformat(timespec="seconds")

    for job_id in finished_job_ids:
        shutil.rmtree(UPLOAD_ROOT / job_id, ignore_errors=True)


def set_job_status(job_id, status, **extra):
    with JOBS_LOCK:
        job = JOBS.get(job_id)
        if not job:
            return
        job["status"] = status
        job.update(extra)


def result_file_entries(result_dir):
    files = sorted(result_dir.glob("*.xlsx"), key=lambda item: item.name)
    return [
        {
            "id": str(index),
            "filename": file_path.name,
            "path": str(file_path),
            "downloaded": False,
        }
        for index, file_path in enumerate(files)
    ]


def register_result_files(job_id, result_dir):
    entries = result_file_entries(result_dir)
    if not entries:
        return
    with JOBS_LOCK:
        job = JOBS.get(job_id)
        if not job:
            return
        job["download_available"] = True
        job["result_files"] = entries


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


def montar_resumo(job, n_anexos=0, saudacao=False):
    progress = job.get("progress") or {}
    is_consulta = job.get("mode") == "consulta"
    titulo = {
        "completed": "✅ *Automação concluída*",
        "stopped": "⏸️ *Automação parada* (resultados salvos)",
        "failed": "❌ *Automação finalizada com erro*",
    }.get(job.get("status"), "*Automação finalizada*")
    rotulo_sucesso = "telefones encontrados" if is_consulta else "leads cadastrados"
    rotulo_pendente = "não encontrados" if is_consulta else "duplicados/não cadastrados"

    total = progress.get("total", 0) or 0
    processados = progress.get("processados", 0) or 0
    sucessos = progress.get("sucessos", 0) or 0
    pendentes = progress.get("pendentes", 0) or 0
    erros = progress.get("erros", 0) or 0
    base = processados or total
    taxa = round((sucessos / base) * 100) if base else 0

    try:
        hora = datetime.fromisoformat(job.get("finished_at")).strftime("%H:%M")
    except (TypeError, ValueError):
        hora = ""

    linha = "━━━━━━━━━━━━━━━"
    if saudacao:
        # Versao amigavel pro grupo do Marketing: apresentacao no lugar do titulo tecnico.
        cabecalho = [
            "👋 Olá! Tudo bem? Aqui é a *automação de importação de telefones e cadastro de leads*. 🤖",
            "",
            "Terminei de processar a planilha que me foi enviada — segue abaixo o resumo e as planilhas em anexo. 📎",
            "",
        ]
    else:
        cabecalho = [titulo]
    partes = cabecalho + [
        linha,
        f"📄 {job.get('filename', '')}",
        f"⚙️ {'Consulta' if is_consulta else 'Cadastro'}   ·   ⏱️ {_formatar_duracao(job.get('started_at'), job.get('finished_at'))}",
        f"📊 {processados}/{total} processados",
        "",
        "*Resultado*",
        f"✅ {sucessos} {rotulo_sucesso}",
        f"➖ {pendentes} {rotulo_pendente}",
        f"⚠️ {erros} erro(s)",
        linha,
        f"🎯 Taxa de sucesso: *{taxa}%*",
    ]
    if n_anexos:
        partes.append(f"📎 {n_anexos} planilha(s) em anexo")
    if hora:
        partes.append(f"🕐 Concluído às {hora}")
    if saudacao:
        partes.append("")
        partes.append("Qualquer coisa, é só chamar o Erick. 🙂")
    return "\n".join(partes)


def notificar_whatsapp(job_id, result_dir):
    """Envia o resumo + planilhas pelo WhatsApp ao fim do job (se configurado)."""
    if not whatsapp.credenciais_ok():
        return
    with JOBS_LOCK:
        job = JOBS.get(job_id)
        snapshot = dict(job) if job else None
    if not snapshot or snapshot.get("status") not in {"completed", "stopped", "failed"}:
        return
    arquivos = [entry["path"] for entry in result_file_entries(result_dir)]
    anexaveis = [a for a in arquivos if Path(a).exists() and Path(a).stat().st_size > 0]
    texto = montar_resumo(snapshot, n_anexos=len(anexaveis))
    texto_grupo = montar_resumo(snapshot, n_anexos=len(anexaveis), saudacao=True)
    # Destinos escolhidos na tela; se vazio (ex.: JS falhou), cai no .env (seguro).
    destinos = snapshot.get("destinos") or None
    ok, msg = whatsapp.notificar(texto, arquivos, texto_grupo=texto_grupo, destinos=destinos)
    append_log(job_id, f"\n{msg}\n")


def run_automation(job_id, excel_path, result_dir, progress_file, stop_file, log_path, mode, sigavi_login, sigavi_senha, headless):
    secrets_to_hide = [sigavi_login, sigavi_senha]
    command = [
        sys.executable,
        "-u",
        str(BASE_DIR / "confio.py"),
        "--excel",
        str(excel_path),
        "--result-dir",
        str(result_dir),
        "--progress-file",
        str(progress_file),
        "--stop-file",
        str(stop_file),
        "--mode",
        mode,
    ]
    if headless:
        command.append("--headless")

    env = os.environ.copy()
    env["SIGAVI_LOGIN"] = sigavi_login
    env["SIGAVI_SENHA"] = sigavi_senha
    env["PYTHONIOENCODING"] = "utf-8"

    set_job_status(job_id, "running", started_at=datetime.now().isoformat(timespec="seconds"))
    append_log(job_id, "Automacao iniciada.\n", secrets_to_hide)
    append_file_log(log_path, "Automacao iniciada.\n", secrets_to_hide)

    return_code = None
    try:
        with RUN_LOCK:
            process = subprocess.Popen(
                command,
                cwd=BASE_DIR,
                env=env,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                encoding="utf-8",
                errors="replace",
                bufsize=1,
            )
            set_job_status(job_id, "running", pid=process.pid)

            if process.stdout:
                for line in process.stdout:
                    if line.startswith("PROGRESS="):
                        update_progress(job_id, line)
                        continue
                    append_log(job_id, line, secrets_to_hide)
                    append_file_log(log_path, line, secrets_to_hide)

            return_code = process.wait()

        register_result_files(job_id, result_dir)

        finished_at = datetime.now().isoformat(timespec="seconds")
        if return_code == 0:
            if stop_file.exists():
                append_log(job_id, "\nAutomacao parada com resultados salvos.\n", secrets_to_hide)
                append_file_log(log_path, "\nAutomacao parada com resultados salvos.\n", secrets_to_hide)
                set_job_status(job_id, "stopped", return_code=return_code, finished_at=finished_at)
            else:
                append_log(job_id, "\nAutomacao concluida.\n", secrets_to_hide)
                append_file_log(log_path, "\nAutomacao concluida.\n", secrets_to_hide)
                set_job_status(job_id, "completed", return_code=return_code, finished_at=finished_at)
        else:
            append_log(job_id, f"\nAutomacao finalizada com erro. Codigo: {return_code}\n", secrets_to_hide)
            append_file_log(log_path, f"\nAutomacao finalizada com erro. Codigo: {return_code}\n", secrets_to_hide)
            set_job_status(job_id, "failed", return_code=return_code, finished_at=finished_at)
    except Exception as exc:
        append_log(job_id, f"\nErro ao executar automacao: {exc}\n", secrets_to_hide)
        append_file_log(log_path, f"\nErro ao executar automacao: {exc}\n", secrets_to_hide)
        set_job_status(
            job_id,
            "failed",
            return_code=return_code,
            finished_at=datetime.now().isoformat(timespec="seconds"),
        )
    finally:
        notificar_whatsapp(job_id, result_dir)
        try:
            excel_path.unlink(missing_ok=True)
        except OSError:
            pass
        try:
            stop_file.unlink(missing_ok=True)
        except OSError:
            pass

        if result_file_entries(result_dir):
            register_result_files(job_id, result_dir)
        else:
            shutil.rmtree(excel_path.parent, ignore_errors=True)


def _iniciar_job(excel_bytes, display_name, mode, sigavi_login, sigavi_senha, headless, destinos=None, extra_logs=None):
    """Cria o job, grava a planilha + backup e dispara a thread de execucao.

    Compartilhado por /jobs (upload novo) e /jobs/<id>/reprocess (so as linhas
    que falharam). Recebe os bytes do .xlsx ja prontos para servir de entrada.
    """
    secure_name = secure_filename(display_name) or "planilha.xlsx"
    job_id = uuid.uuid4().hex
    job_dir = UPLOAD_ROOT / job_id
    job_dir.mkdir(parents=True, exist_ok=False)
    backup_dir = BACKUP_ROOT / backup_slug(secure_name)
    backup_dir.mkdir(parents=True, exist_ok=False)
    result_dir = backup_dir / "resultados"
    result_dir.mkdir(exist_ok=False)
    progress_file = backup_dir / "progresso.json"
    stop_file = job_dir / "parar-e-salvar.flag"
    log_path = backup_dir / "log.txt"
    # O job_dir ja e unico (UPLOAD_ROOT/job_id), entao o arquivo dentro dele nao
    # precisa do prefixo do job_id. Sem ele, o nome do resultado sai limpo
    # (ex.: resultado_New_Time_encontrados.xlsx em vez de resultado_<hash>_New_Time_...).
    excel_path = job_dir / secure_name
    excel_path.write_bytes(excel_bytes)
    shutil.copy2(excel_path, backup_dir / f"entrada_{secure_name}")

    logs = [
        f"Modo selecionado: {'Somente consulta' if mode == 'consulta' else 'Somente cadastro'}\n",
        f"Planilha recebida: {display_name}\n",
        f"Backup criado em: {backup_dir}\n",
    ]
    if extra_logs:
        logs = list(extra_logs) + logs

    with JOBS_LOCK:
        JOBS[job_id] = {
            "id": job_id,
            "status": "queued",
            "mode": mode,
            "filename": display_name,
            "created_at": datetime.now().isoformat(timespec="seconds"),
            "started_at": None,
            "finished_at": None,
            "return_code": None,
            "pid": None,
            "backup_dir": str(backup_dir),
            "stop_file": str(stop_file),
            "destinos": list(destinos or []),
            "download_available": False,
            "result_files": [],
            "result_deleted_at": None,
            "progress": None,
            "logs": logs,
        }

    thread = threading.Thread(
        target=run_automation,
        args=(job_id, excel_path, result_dir, progress_file, stop_file, log_path, mode, sigavi_login, sigavi_senha, headless),
        daemon=True,
    )
    thread.start()
    return job_id


@app.get("/")
def index():
    return render_template(
        "index.html",
        authenticated=session.get("authenticated", False),
        login_error=request.args.get("login_error"),
    )


@app.get("/destinos")
@login_required
def listar_destinos():
    """Lista curada de destinos pro front montar os checkboxes 'Enviar para'."""
    return jsonify({"destinos": carregar_destinos()})


@app.post("/login")
def login():
    if not check_csrf():
        return redirect(url_for("index", login_error="Sessao expirada. Tente novamente."))

    if not secrets.compare_digest(request.form.get("password", ""), APP_PASSWORD):
        return redirect(url_for("index", login_error="Senha de acesso invalida."))

    session["authenticated"] = True
    session["csrf_token"] = secrets.token_urlsafe(32)
    return redirect(url_for("index"))


@app.post("/logout")
@login_required
def logout():
    if not check_csrf():
        return redirect(url_for("index", login_error="Sessao expirada. Tente novamente."))
    session.clear()
    return redirect(url_for("index"))


@app.post("/jobs")
@login_required
def create_job():
    if not check_csrf():
        return jsonify({"error": "Sessao expirada. Atualize a pagina."}), 400

    if has_active_job():
        return jsonify({"error": "Ja existe uma automacao em andamento."}), 409

    cleanup_finished_job_files()

    sigavi_login = request.form.get("sigavi_login", "").strip()
    sigavi_senha = request.form.get("sigavi_senha", "")
    upload = request.files.get("planilha")
    headless = request.form.get("headless") == "on"
    mode = request.form.get("mode", "consulta")
    destinos = validar_destinos(request.form.getlist("destinos"))

    if not sigavi_login or not sigavi_senha:
        return jsonify({"error": "Informe login e senha do Sigavi."}), 400
    if mode not in {"consulta", "cadastro"}:
        return jsonify({"error": "Modo de execucao invalido."}), 400
    if not upload or not upload.filename:
        return jsonify({"error": "Envie uma planilha .xlsx."}), 400

    original_name = secure_filename(upload.filename)
    extension = Path(original_name).suffix.lower()
    if extension != ".xlsx":
        return jsonify({"error": "A planilha precisa estar no formato .xlsx."}), 400

    try:
        df, _sheet = ler_planilha_upload(upload)
        validos, _email_col = dados_validos_planilha(df, mode)
    except Exception as exc:
        return jsonify({"error": f"Nao foi possivel ler a planilha: {exc}"}), 400
    if validos == 0:
        requisito = "e-mail" if mode == "consulta" else "telefone valido"
        return jsonify({"error": f"A planilha nao possui nenhuma linha com {requisito}."}), 400

    upload.stream.seek(0)
    excel_bytes = upload.read()
    job_id = _iniciar_job(excel_bytes, upload.filename, mode, sigavi_login, sigavi_senha, headless, destinos=destinos)

    return jsonify({"job_id": job_id})


@app.get("/jobs/<job_id>")
@login_required
def get_job(job_id):
    with JOBS_LOCK:
        job = JOBS.get(job_id)
        if not job:
            return jsonify({"error": "Execucao nao encontrada."}), 404
        payload = dict(job)
        payload["logs"] = list(job["logs"])
        result_files = [
            {"id": item["id"], "filename": item["filename"]}
            for item in job.get("result_files", [])
            if not item.get("downloaded")
        ]
        payload["result_files"] = result_files
        payload["download_available"] = bool(result_files)
    return jsonify(payload)


@app.post("/jobs/<job_id>/stop")
@login_required
def stop_job(job_id):
    if not check_csrf():
        return jsonify({"error": "Sessao expirada. Atualize a pagina."}), 400

    with JOBS_LOCK:
        job = JOBS.get(job_id)
        if not job:
            return jsonify({"error": "Execucao nao encontrada."}), 404
        if job["status"] not in RUNNING_STATUSES:
            return jsonify({"error": "Execucao nao esta rodando."}), 409
        stop_file = Path(job["stop_file"])
        job["status"] = "stopping"

    stop_file.parent.mkdir(parents=True, exist_ok=True)
    stop_file.write_text("parar", encoding="utf-8")
    append_log(job_id, "\nParada solicitada. Salvando resultados no proximo ponto seguro...\n")
    return jsonify({"ok": True})


def _serve_result_file(job_id, file_id=None):
    selected_file = None
    remaining_files = []
    with JOBS_LOCK:
        job = JOBS.get(job_id)
        if not job:
            return jsonify({"error": "Resultado nao disponivel."}), 404

        available_files = [
            item
            for item in job.get("result_files", [])
            if not item.get("downloaded")
        ]
        if file_id is None and available_files:
            selected_file = available_files[0]
        else:
            selected_file = next((item for item in available_files if item["id"] == file_id), None)

        if not selected_file:
            return jsonify({"error": "Resultado nao disponivel."}), 404

        result_path = Path(selected_file["path"])
        result_filename = selected_file["filename"]

    if not result_path.exists():
        return jsonify({"error": "Arquivo de resultado nao encontrado."}), 404

    data = result_path.read_bytes()

    headers = {"Content-Disposition": f'attachment; filename="{secure_filename(result_filename)}"'}
    return Response(data, mimetype=XLSX_MIMETYPE, headers=headers)


@app.get("/jobs/<job_id>/download")
@login_required
def download_result(job_id):
    return _serve_result_file(job_id)


@app.get("/jobs/<job_id>/download/<file_id>")
@login_required
def download_result_file(job_id, file_id):
    return _serve_result_file(job_id, file_id)


@app.post("/jobs/<job_id>/reprocess")
@login_required
def reprocess_errors(job_id):
    """Roda de novo apenas as linhas que falharam (erro_consulta/erro_cadastro).

    Le o progresso salvo do job original, filtra a planilha de entrada (do backup)
    pelas linhas com erro e inicia um novo job so com elas, no mesmo modo.
    """
    import pandas as pd

    if not check_csrf():
        return jsonify({"error": "Sessao expirada. Atualize a pagina."}), 400
    if has_active_job():
        return jsonify({"error": "Ja existe uma automacao em andamento."}), 409

    with JOBS_LOCK:
        job = JOBS.get(job_id)
        snapshot = dict(job) if job else None
    if not snapshot:
        return jsonify({"error": "Execucao original nao encontrada."}), 404
    if snapshot["status"] in RUNNING_STATUSES:
        return jsonify({"error": "A execucao ainda esta rodando."}), 409

    sigavi_login = request.form.get("sigavi_login", "").strip()
    sigavi_senha = request.form.get("sigavi_senha", "")
    headless = request.form.get("headless") == "on"
    if not sigavi_login or not sigavi_senha:
        return jsonify({"error": "Informe login e senha do Sigavi para reprocessar."}), 400

    mode = snapshot.get("mode", "consulta")
    status_erro = "erro_consulta" if mode == "consulta" else "erro_cadastro"
    chave_resultados = "resultados_email" if mode == "consulta" else "resultados_cadastro"

    backup_dir = Path(snapshot.get("backup_dir", ""))
    progresso_path = backup_dir / "progresso.json"
    if not progresso_path.exists():
        return jsonify({"error": "Nao ha dados de progresso para reprocessar."}), 400

    try:
        dados = json.loads(progresso_path.read_text(encoding="utf-8"))
    except (ValueError, OSError) as exc:
        return jsonify({"error": f"Nao foi possivel ler o progresso: {exc}"}), 400

    linhas_erro = sorted(
        {
            r.get("Linha")
            for r in dados.get(chave_resultados, [])
            if r.get("Status") == status_erro and isinstance(r.get("Linha"), int)
        }
    )
    if not linhas_erro:
        return jsonify({"error": "Nenhum erro para reprocessar nesta execucao."}), 400

    entradas = sorted(backup_dir.glob("entrada_*.xlsx"))
    if not entradas:
        return jsonify({"error": "Planilha de entrada do job original nao encontrada."}), 400

    try:
        excel_file = pd.ExcelFile(entradas[0])
        sheet = excel_file.sheet_names[0]
        df = excel_file.parse(sheet)
    except Exception as exc:
        return jsonify({"error": f"Nao foi possivel ler a planilha de entrada: {exc}"}), 400

    # 'Linha' = index + 1 no confio.py, entao a posicao 0-based e (Linha - 1).
    posicoes = [linha - 1 for linha in linhas_erro if 0 <= linha - 1 < len(df)]
    if not posicoes:
        return jsonify({"error": "As linhas com erro nao batem com a planilha original."}), 400

    buffer = io.BytesIO()
    df.iloc[posicoes].to_excel(buffer, index=False, sheet_name=str(sheet)[:31])
    excel_bytes = buffer.getvalue()

    display_name = f"reprocesso_{snapshot.get('filename', 'planilha.xlsx')}"
    extra_logs = [f"Reprocessando {len(posicoes)} linha(s) com erro do job anterior.\n"]
    novo_job_id = _iniciar_job(
        excel_bytes, display_name, mode, sigavi_login, sigavi_senha, headless,
        destinos=snapshot.get("destinos"), extra_logs=extra_logs
    )

    return jsonify({"job_id": novo_job_id, "reprocessadas": len(posicoes)})


@app.post("/preview")
@login_required
def preview_planilha():
    """Le a planilha enviada e devolve uma previa + contagens (sem executar nada).

    Espelha a leitura do confio.py (primeira aba + renomeacao de colunas) para que
    o total/validos/ignorados mostrados na tela batam com o que a automacao vai usar.
    """
    if not check_csrf():
        return jsonify({"error": "Sessao expirada. Atualize a pagina."}), 400

    upload = request.files.get("planilha")
    if not upload or not upload.filename:
        return jsonify({"error": "Nenhum arquivo enviado."}), 400
    if Path(secure_filename(upload.filename)).suffix.lower() != ".xlsx":
        return jsonify({"error": "Formato invalido. Use .xlsx."}), 400

    mode = request.form.get("mode", "consulta")
    if mode not in {"consulta", "cadastro"}:
        return jsonify({"error": "Modo de execucao invalido."}), 400
    try:
        import pandas as pd

        df, sheet = ler_planilha_upload(upload)
    except Exception as exc:  # planilha corrompida, vazia, etc.
        return jsonify({"error": f"Nao foi possivel ler a planilha: {exc}"}), 400

    total = int(len(df))
    validos, email_col = dados_validos_planilha(df, mode)
    ignorados = max(0, total - validos)

    # Colunas amigaveis para a previa (so as que existirem na planilha).
    empreend_col = next((c for c in df.columns if "EMPREEND" in str(c).upper()), None)
    fac_col = _detectar_coluna(df, RE_FAC_COL, full=True)
    nome_col = _detectar_coluna(df, RE_NOME_COL) or ("NOME" if "NOME" in df.columns else None)
    mapa_exibicao = [
        ("FAC", fac_col),
        ("Nome", nome_col),
        ("Email", email_col),
        ("Telefone", "FONE2"),
        ("Corretor", "CORRETOR DE ORIGEM"),
        ("Empreendimento", empreend_col),
    ]
    presentes = [(rotulo, col) for rotulo, col in mapa_exibicao if col and col in df.columns]
    tem = {rotulo: (col is not None and col in df.columns) for rotulo, col in mapa_exibicao}

    # Validacao por modo: cada modo precisa de colunas diferentes.
    #  - consulta: basta UM criterio de busca (FAC, email ou nome) — a automacao
    #    busca telefone+email no Sigavi a partir dele, nessa ordem de prioridade;
    #  - cadastro: precisa de Nome, Telefone, Corretor e Empreendimento.
    if mode == "consulta":
        if tem["FAC"] or tem["Email"] or tem["Nome"]:
            faltando = []
            status_previa = "ok"
            criterio = "FAC" if tem["FAC"] else ("e-mail" if tem["Email"] else "nome")
            mensagem_previa = (
                f"Planilha de consulta OK — {validos} linha(s) com {criterio} pra buscar telefone + e-mail no Sigavi."
            )
        else:
            faltando = ["FAC/Email/Nome"]
            status_previa = "aviso"
            mensagem_previa = (
                "Nenhuma coluna de FAC, e-mail ou nome encontrada. No modo Consulta a planilha precisa de pelo menos uma delas."
            )
    else:  # cadastro
        necessarias = ["Nome", "Telefone", "Corretor", "Empreendimento"]
        faltando = [c for c in necessarias if not tem[c]]
        if not faltando:
            status_previa = "ok"
            mensagem_previa = "Planilha de cadastro OK — todas as colunas necessárias foram encontradas."
        else:
            status_previa = "aviso"
            mensagem_previa = (
                f"Faltam colunas pro cadastro: {', '.join(faltando)}. A automação usa o que houver disponível."
            )

    linhas = []
    for _, row in df.head(5).iterrows():
        linhas.append({rotulo: ("" if pd.isna(row.get(col)) else str(row.get(col))) for rotulo, col in presentes})

    return jsonify({
        "sheet": str(sheet),
        "total": total,
        "validos": validos,
        "ignorados": ignorados,
        "colunas": [rotulo for rotulo, _ in presentes],
        "rows": linhas,
        "faltando": faltando,
        "status_previa": status_previa,
        "mensagem_previa": mensagem_previa,
        "mode": mode,
    })


@app.get("/health")
def health():
    return jsonify({"ok": True})


if __name__ == "__main__":
    host = os.getenv("FLASK_HOST", "127.0.0.1")
    port = int(os.getenv("FLASK_PORT", "5000"))
    app.run(host=host, port=port, debug=False)
