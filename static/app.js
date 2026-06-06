const form = document.querySelector("#run-form");
const submitButton = document.querySelector("#submit-button");
const stopButton = document.querySelector("#stop-button");
const statusBadge = document.querySelector("#status-badge");
const jobIdLabel = document.querySelector("#job-id");
const logOutput = document.querySelector("#log-output");
const downloadCard = document.querySelector("#download-card");
const downloadList = document.querySelector("#download-list");
const metricProcessed = document.querySelector("#metric-processed");
const metricSuccess = document.querySelector("#metric-success");
const metricPending = document.querySelector("#metric-pending");
const metricErrors = document.querySelector("#metric-errors");
const progressBar = document.querySelector("#progress-bar");
const progressLabel = document.querySelector("#progress-label");

let currentJobId = null;
let pollTimer = null;

function setStatus(label, state) {
  statusBadge.textContent = label;
  statusBadge.className = `status-badge ${state}`;
}

function setBusy(isBusy) {
  submitButton.disabled = isBusy;
  stopButton.disabled = !isBusy;
  submitButton.textContent = isBusy ? "Executando..." : "Executar automação";
}

function renderLogs(lines) {
  logOutput.textContent = lines.length ? lines.join("") : "Sem logs ainda.";
  logOutput.scrollTop = logOutput.scrollHeight;
}

function setProgressBar(processed, total) {
  const pct = total > 0 ? Math.min(100, Math.round((processed / total) * 100)) : 0;
  if (progressBar) {
    progressBar.style.width = `${pct}%`;
  }
  if (progressLabel) {
    progressLabel.textContent = total > 0 ? `${processed}/${total} (${pct}%)` : "";
  }
}

function updateMetrics(data) {
  // Fonte de verdade: contadores reais enviados pelo backend (campo progress).
  const progress = data.progress;
  if (progress) {
    metricProcessed.textContent = `${progress.processados}/${progress.total}`;
    metricSuccess.textContent = progress.sucessos;
    metricPending.textContent = progress.pendentes;
    metricErrors.textContent = progress.erros;
    setProgressBar(progress.processados, progress.total);
    return;
  }

  // Fallback (antes do primeiro progresso chegar): contagem aproximada por log.
  const text = (data.logs || []).join("");
  metricProcessed.textContent = String((text.match(/\[CADASTRADO\]|\[DUPLICADO\]|\[NAO CADASTRADO\]/g) || []).length);
  metricSuccess.textContent = (text.match(/\[✓\]|\[CADASTRADO\]/g) || []).length;
  metricPending.textContent = (text.match(/\[✗\]|\[DUPLICADO\]|\[NAO CADASTRADO\]/g) || []).length;
  metricErrors.textContent = (text.match(/\[ERRO\]|erro_consulta|erro_cadastro/g) || []).length;
}

function hideDownload() {
  downloadCard.classList.add("hidden");
  downloadList.innerHTML = "";
}

function updateDownload(data) {
  const files = data.result_files || [];
  if (!data.download_available || files.length === 0 || !currentJobId) {
    hideDownload();
    return;
  }

  downloadList.innerHTML = "";
  for (const file of files) {
    const link = document.createElement("a");
    link.className = "download-button";
    link.href = `/jobs/${currentJobId}/download/${file.id}`;
    link.textContent = file.filename;
    downloadList.appendChild(link);
  }
  downloadCard.classList.remove("hidden");
}

async function pollJob() {
  if (!currentJobId) {
    return;
  }

  const response = await fetch(`/jobs/${currentJobId}`);
  const data = await response.json();

  if (!response.ok) {
    setStatus("Erro", "error");
    renderLogs([data.error || "Nao foi possivel buscar a execucao."]);
    setBusy(false);
    return;
  }

  renderLogs(data.logs || []);
  updateMetrics(data);
  updateDownload(data);

  if (data.status === "queued") {
    setStatus("Na fila", "running");
  } else if (data.status === "running") {
    setStatus("Executando", "running");
  } else if (data.status === "stopping") {
    setStatus("Parando", "running");
  } else if (data.status === "completed") {
    setStatus("Concluido", "done");
    setBusy(false);
    return;
  } else if (data.status === "stopped") {
    setStatus("Salvo", "done");
    setBusy(false);
    return;
  } else {
    setStatus("Erro", "error");
    setBusy(false);
    return;
  }

  pollTimer = window.setTimeout(pollJob, 1200);
}

form.addEventListener("submit", async (event) => {
  event.preventDefault();
  window.clearTimeout(pollTimer);
  hideDownload();

  setBusy(true);
  setStatus("Enviando", "running");
  jobIdLabel.textContent = "Preparando";
  logOutput.textContent = "Enviando planilha...\n";

  try {
    const response = await fetch("/jobs", {
      method: "POST",
      body: new FormData(form),
    });
    const data = await response.json();

    if (!response.ok) {
      throw new Error(data.error || "Falha ao iniciar automacao.");
    }

    currentJobId = data.job_id;
    jobIdLabel.textContent = currentJobId.slice(0, 8);
    setStatus("Executando", "running");
    pollJob();
  } catch (error) {
    setStatus("Erro", "error");
    logOutput.textContent = `${error.message}\n`;
    setBusy(false);
  }
});

downloadList.addEventListener("click", (event) => {
  if (!event.target.matches("a")) {
    return;
  }

  event.target.textContent = "Baixando...";
  window.setTimeout(pollJob, 1000);
});

stopButton.addEventListener("click", async () => {
  if (!currentJobId || stopButton.disabled) {
    return;
  }

  stopButton.disabled = true;
  setStatus("Parando", "running");

  const body = new FormData();
  body.append("csrf_token", form.querySelector("[name='csrf_token']").value);

  try {
    const response = await fetch(`/jobs/${currentJobId}/stop`, {
      method: "POST",
      body,
    });
    const data = await response.json();
    if (!response.ok) {
      throw new Error(data.error || "Nao foi possivel parar a execucao.");
    }
    logOutput.textContent += "\nParada solicitada. Salvando no proximo ponto seguro...\n";
    pollJob();
  } catch (error) {
    setStatus("Erro", "error");
    logOutput.textContent += `\n${error.message}\n`;
    stopButton.disabled = false;
  }
});
