import os
import secrets
import shutil
import subprocess
import sys
import threading
import uuid
from datetime import datetime
from functools import wraps
from pathlib import Path

from dotenv import load_dotenv
from flask import Flask, Response, jsonify, redirect, render_template, request, session, url_for
from werkzeug.utils import secure_filename


BASE_DIR = Path(__file__).resolve().parent
UPLOAD_ROOT = BASE_DIR / "uploads"
MAX_LOG_LINES = 2000
RUNNING_STATUSES = {"queued", "running"}
XLSX_MIMETYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

load_dotenv(BASE_DIR / ".env")
UPLOAD_ROOT.mkdir(exist_ok=True)

app = Flask(__name__)
app.secret_key = os.getenv("APP_SECRET_KEY") or secrets.token_hex(32)
app.config["MAX_CONTENT_LENGTH"] = int(os.getenv("MAX_UPLOAD_MB", "25")) * 1024 * 1024
app.config.update(
    SESSION_COOKIE_HTTPONLY=True,
    SESSION_COOKIE_SAMESITE="Lax",
)

APP_PASSWORD_CONFIGURED = bool(os.getenv("APP_PASSWORD"))
APP_PASSWORD = os.getenv("APP_PASSWORD", "marketing123")
JOBS = {}
JOBS_LOCK = threading.Lock()
RUN_LOCK = threading.Lock()

if not APP_PASSWORD_CONFIGURED:
    print("AVISO: APP_PASSWORD nao configurado. Senha local padrao: marketing123")


def get_csrf_token():
    if "csrf_token" not in session:
        session["csrf_token"] = secrets.token_urlsafe(32)
    return session["csrf_token"]


@app.context_processor
def inject_template_values():
    return {"csrf_token": get_csrf_token()}


def check_csrf():
    return request.form.get("csrf_token") == session.get("csrf_token")


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
            job["result_path"] = None
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


def latest_result_file(result_dir):
    files = list(result_dir.glob("*.xlsx"))
    if not files:
        return None
    return max(files, key=lambda item: item.stat().st_mtime)


def register_result_file(job_id, result_path):
    if not result_path:
        return
    with JOBS_LOCK:
        job = JOBS.get(job_id)
        if not job:
            return
        job["download_available"] = True
        job["result_filename"] = result_path.name
        job["result_path"] = str(result_path)


def run_automation(job_id, excel_path, result_dir, sigavi_login, sigavi_senha, headless):
    secrets_to_hide = [sigavi_login, sigavi_senha]
    command = [
        sys.executable,
        "-u",
        str(BASE_DIR / "confio.py"),
        "--excel",
        str(excel_path),
        "--result-dir",
        str(result_dir),
    ]
    if headless:
        command.append("--headless")

    env = os.environ.copy()
    env["SIGAVI_LOGIN"] = sigavi_login
    env["SIGAVI_SENHA"] = sigavi_senha
    env["PYTHONIOENCODING"] = "utf-8"

    set_job_status(job_id, "running", started_at=datetime.now().isoformat(timespec="seconds"))
    append_log(job_id, "Automacao iniciada.\n", secrets_to_hide)

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
                    append_log(job_id, line, secrets_to_hide)

            return_code = process.wait()

        result_path = latest_result_file(result_dir)
        if result_path:
            register_result_file(job_id, result_path)

        finished_at = datetime.now().isoformat(timespec="seconds")
        if return_code == 0:
            append_log(job_id, "\nAutomacao concluida.\n", secrets_to_hide)
            set_job_status(job_id, "completed", return_code=return_code, finished_at=finished_at)
        else:
            append_log(job_id, f"\nAutomacao finalizada com erro. Codigo: {return_code}\n", secrets_to_hide)
            set_job_status(job_id, "failed", return_code=return_code, finished_at=finished_at)
    except Exception as exc:
        append_log(job_id, f"\nErro ao executar automacao: {exc}\n", secrets_to_hide)
        set_job_status(
            job_id,
            "failed",
            return_code=return_code,
            finished_at=datetime.now().isoformat(timespec="seconds"),
        )
    finally:
        try:
            excel_path.unlink(missing_ok=True)
        except OSError:
            pass

        result_path = latest_result_file(result_dir)
        if result_path:
            register_result_file(job_id, result_path)
        else:
            shutil.rmtree(excel_path.parent, ignore_errors=True)


@app.get("/")
def index():
    return render_template(
        "index.html",
        authenticated=session.get("authenticated", False),
        login_error=request.args.get("login_error"),
    )


@app.post("/login")
def login():
    if not check_csrf():
        return redirect(url_for("index", login_error="Sessao expirada. Tente novamente."))

    if request.form.get("password") != APP_PASSWORD:
        return redirect(url_for("index", login_error="Senha de acesso invalida."))

    session["authenticated"] = True
    session["csrf_token"] = secrets.token_urlsafe(32)
    return redirect(url_for("index"))


@app.post("/logout")
@login_required
def logout():
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

    if not sigavi_login or not sigavi_senha:
        return jsonify({"error": "Informe login e senha do Sigavi."}), 400
    if not upload or not upload.filename:
        return jsonify({"error": "Envie uma planilha .xlsx."}), 400

    original_name = secure_filename(upload.filename)
    extension = Path(original_name).suffix.lower()
    if extension != ".xlsx":
        return jsonify({"error": "A planilha precisa estar no formato .xlsx."}), 400

    job_id = uuid.uuid4().hex
    job_dir = UPLOAD_ROOT / job_id
    job_dir.mkdir(parents=True, exist_ok=False)
    result_dir = job_dir / "resultado"
    result_dir.mkdir(exist_ok=False)
    excel_path = job_dir / f"{job_id}_{original_name}"
    upload.save(excel_path)

    with JOBS_LOCK:
        JOBS[job_id] = {
            "id": job_id,
            "status": "queued",
            "filename": upload.filename,
            "created_at": datetime.now().isoformat(timespec="seconds"),
            "started_at": None,
            "finished_at": None,
            "return_code": None,
            "pid": None,
            "download_available": False,
            "result_filename": None,
            "result_path": None,
            "result_deleted_at": None,
            "logs": [f"Planilha recebida: {upload.filename}\n"],
        }

    thread = threading.Thread(
        target=run_automation,
        args=(job_id, excel_path, result_dir, sigavi_login, sigavi_senha, headless),
        daemon=True,
    )
    thread.start()

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
    return jsonify(payload)


@app.get("/jobs/<job_id>/download")
@login_required
def download_result(job_id):
    with JOBS_LOCK:
        job = JOBS.get(job_id)
        if not job or not job.get("download_available") or not job.get("result_path"):
            return jsonify({"error": "Resultado nao disponivel."}), 404
        result_path = Path(job["result_path"])
        result_filename = job.get("result_filename") or "resultado.xlsx"

    if not result_path.exists():
        return jsonify({"error": "Arquivo de resultado nao encontrado."}), 404

    data = result_path.read_bytes()
    shutil.rmtree(UPLOAD_ROOT / job_id, ignore_errors=True)

    with JOBS_LOCK:
        job = JOBS.get(job_id)
        if job:
            job["download_available"] = False
            job["result_path"] = None
            job["result_deleted_at"] = datetime.now().isoformat(timespec="seconds")

    headers = {"Content-Disposition": f'attachment; filename="{secure_filename(result_filename)}"'}
    return Response(data, mimetype=XLSX_MIMETYPE, headers=headers)


@app.get("/health")
def health():
    return jsonify({"ok": True})


if __name__ == "__main__":
    host = os.getenv("FLASK_HOST", "127.0.0.1")
    port = int(os.getenv("FLASK_PORT", "5000"))
    app.run(host=host, port=port, debug=False)
