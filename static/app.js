const form = document.querySelector("#run-form");
const submitButton = document.querySelector("#submit-button");
const statusBadge = document.querySelector("#status-badge");
const jobIdLabel = document.querySelector("#job-id");
const logOutput = document.querySelector("#log-output");
const downloadCard = document.querySelector("#download-card");
const downloadLink = document.querySelector("#download-link");
const downloadFilename = document.querySelector("#download-filename");

let currentJobId = null;
let pollTimer = null;

function setStatus(label, state) {
  statusBadge.textContent = label;
  statusBadge.className = `status-badge ${state}`;
}

function setBusy(isBusy) {
  submitButton.disabled = isBusy;
  submitButton.textContent = isBusy ? "Executando..." : "Executar automação";
}

function renderLogs(lines) {
  logOutput.textContent = lines.length ? lines.join("") : "Sem logs ainda.";
  logOutput.scrollTop = logOutput.scrollHeight;
}

function hideDownload() {
  downloadCard.classList.add("hidden");
  downloadLink.removeAttribute("href");
  downloadFilename.textContent = "resultado.xlsx";
}

function updateDownload(data) {
  if (!data.download_available || !data.result_filename || !currentJobId) {
    hideDownload();
    return;
  }

  downloadFilename.textContent = data.result_filename;
  downloadLink.href = `/jobs/${currentJobId}/download`;
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
  } else if (data.status === "completed") {
    setStatus("Concluido", "done");
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

downloadLink.addEventListener("click", () => {
  window.setTimeout(hideDownload, 800);
});
