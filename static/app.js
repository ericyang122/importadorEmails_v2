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
  updateMetrics(lines);
}

function updateMetrics(lines) {
  const text = lines.join("");
  const emailMatches = [...text.matchAll(/\[Email\s+(\d+)\/(\d+)\]/g)];
  const lastEmail = emailMatches.at(-1);
  const processed = lastEmail ? `${lastEmail[1]}/${lastEmail[2]}` : String((text.match(/\[CADASTRADO\]|\[DUPLICADO\]|\[NAO CADASTRADO\]/g) || []).length);
  const successes = (text.match(/\[✓\]|\[CADASTRADO\]/g) || []).length;
  const pending = (text.match(/\[✗\]|\[DUPLICADO\]|\[NAO CADASTRADO\]/g) || []).length;
  const errors = (text.match(/\[ERRO\]|erro_consulta|erro_cadastro/g) || []).length;

  metricProcessed.textContent = processed;
  metricSuccess.textContent = successes;
  metricPending.textContent = pending;
  metricErrors.textContent = errors;
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
