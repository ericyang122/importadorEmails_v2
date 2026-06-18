// ===================== REFERÊNCIAS =====================
const form = document.querySelector("#run-form");
const submitButton = document.querySelector("#submit-button");
const stopButton = document.querySelector("#stop-button");
const clearButton = document.querySelector("#clear-button");
const statusBadge = document.querySelector("#status-badge");
const jobIdLabel = document.querySelector("#job-id");
const logOutput = document.querySelector("#log-output");

const loginInput = form.querySelector("[name='sigavi_login']");
const senhaInput = form.querySelector("[name='sigavi_senha']");

// Upload
const dropzone = document.querySelector("#dropzone");
const fileInput = document.querySelector("#file-input");
const dzEmpty = document.querySelector("#dropzone-empty");
const dzFile = document.querySelector("#dropzone-file");
const fileNameEl = document.querySelector("#file-name");
const fileSizeEl = document.querySelector("#file-size");
const fileClear = document.querySelector("#file-clear");
const fileMessage = document.querySelector("#file-message");

// Resumo
const sumMode = document.querySelector("#sum-mode");
const sumFile = document.querySelector("#sum-file");
const sumTotal = document.querySelector("#sum-total");
const sumValid = document.querySelector("#sum-valid");
const sumIgnored = document.querySelector("#sum-ignored");
const sumStatus = document.querySelector("#sum-status");

// Prévia
const previewBlock = document.querySelector("#preview-block");
const previewTable = document.querySelector("#preview-table");
const previewWarning = document.querySelector("#preview-warning");

// Progresso
const progressBar = document.querySelector("#progress-bar");
const progressPercent = document.querySelector("#progress-percent");
const progressLabel = document.querySelector("#progress-label");
const progressTime = document.querySelector("#progress-time");

// Cards
const metricProcessed = document.querySelector("#metric-processed");
const metricSuccess = document.querySelector("#metric-success");
const metricPending = document.querySelector("#metric-pending");
const metricErrors = document.querySelector("#metric-errors");
const metricIgnored = document.querySelector("#metric-ignored");
const metricSuccessName = document.querySelector("#metric-success-name");
const metricPendingName = document.querySelector("#metric-pending-name");

// Resultado final
const resultPanel = document.querySelector("#result-panel");
const resultIcon = document.querySelector("#result-icon");
const resultTitle = document.querySelector("#result-title");
const resultSubtitle = document.querySelector("#result-subtitle");
const resultTotal = document.querySelector("#result-total");
const resultSuccess = document.querySelector("#result-success");
const resultPending = document.querySelector("#result-pending");
const resultErrors = document.querySelector("#result-errors");
const downloadList = document.querySelector("#download-list");
const reprocessButton = document.querySelector("#reprocess-button");
const reprocessMessage = document.querySelector("#reprocess-message");

// Ações
const actionsFeed = document.querySelector("#actions-feed");

let currentJobId = null;
let reprocessJobId = null;
let pollTimer = null;
let ignoradosPreview = 0;
let previewValido = false;

function csrf() {
  return form.querySelector("[name='csrf_token']").value;
}
function modoAtual() {
  return form.querySelector("[name='mode']:checked").value;
}
function rotuloModo(mode) {
  return mode === "consulta" ? "Buscar telefones" : "Cadastrar leads";
}

// ===================== ESTADO DOS BOTÕES =====================
function isRunning() {
  return ["queued", "running", "stopping"].includes(statusBadge.classList[1]);
}

function refreshSubmitState() {
  const pronto =
    loginInput.value.trim() &&
    senhaInput.value.trim() &&
    fileInput.files.length > 0 &&
    previewValido &&
    !currentJobId;
  submitButton.disabled = !pronto;
}

[loginInput, senhaInput].forEach((el) => el.addEventListener("input", refreshSubmitState));

// ===================== STATUS / BADGE =====================
function setStatus(label, state) {
  statusBadge.textContent = label;
  statusBadge.className = `status-badge ${state}`;
  sumStatus.textContent = label;
}

function setBusy(isBusy) {
  stopButton.disabled = !isBusy;
  clearButton.disabled = isBusy;
  submitButton.textContent = isBusy ? "Executando..." : "▶ Executar automação";
  if (isBusy) submitButton.disabled = true;
  else refreshSubmitState();
}

// ===================== MODO =====================
function aplicarModo() {
  const mode = modoAtual();
  sumMode.textContent = rotuloModo(mode);
  if (mode === "consulta") {
    metricSuccessName.textContent = "Encontrados";
    metricPendingName.textContent = "Não encontrados";
  } else {
    metricSuccessName.textContent = "Cadastrados";
    metricPendingName.textContent = "Duplicados";
  }
}
form.querySelectorAll("[name='mode']").forEach((el) =>
  el.addEventListener("change", () => {
    aplicarModo();
    if (fileInput.files.length > 0) carregarPreview(); // contagens dependem do modo
  })
);

// ===================== UPLOAD =====================
function formatarTamanho(bytes) {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(0)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}

function limparArquivo() {
  fileInput.value = "";
  dzEmpty.classList.remove("hidden");
  dzFile.classList.add("hidden");
  dropzone.classList.remove("has-file");
  fileMessage.textContent = "";
  fileMessage.className = "field-hint";
  previewBlock.classList.add("hidden");
  sumFile.textContent = "—";
  sumTotal.textContent = "—";
  sumValid.textContent = "—";
  sumIgnored.textContent = "—";
  ignoradosPreview = 0;
  previewValido = false;
  refreshSubmitState();
}

function aceitarArquivo(file) {
  const nome = file.name.toLowerCase();
  if (!nome.endsWith(".xlsx")) {
    limparArquivo();
    fileMessage.textContent = "Arquivo inválido. Selecione uma planilha .xlsx.";
    fileMessage.className = "field-hint error";
    return;
  }
  // joga o arquivo no input real (pra ir junto no submit do formulário)
  const dt = new DataTransfer();
  dt.items.add(file);
  fileInput.files = dt.files;

  fileNameEl.textContent = file.name;
  fileSizeEl.textContent = formatarTamanho(file.size);
  dzEmpty.classList.add("hidden");
  dzFile.classList.remove("hidden");
  dropzone.classList.add("has-file");
  fileMessage.textContent = "";
  fileMessage.className = "field-hint";
  sumFile.textContent = file.name;
  previewValido = false;
  refreshSubmitState();
  carregarPreview();
}

dropzone.addEventListener("click", (e) => {
  if (e.target === fileClear) return;
  fileInput.click();
});
dropzone.addEventListener("keydown", (e) => {
  if (e.key === "Enter" || e.key === " ") {
    e.preventDefault();
    fileInput.click();
  }
});
fileInput.addEventListener("change", () => {
  if (fileInput.files.length > 0) aceitarArquivo(fileInput.files[0]);
});
fileClear.addEventListener("click", (e) => {
  e.stopPropagation();
  limparArquivo();
});
["dragenter", "dragover"].forEach((ev) =>
  dropzone.addEventListener(ev, (e) => {
    e.preventDefault();
    dropzone.classList.add("dragover");
  })
);
["dragleave", "drop"].forEach((ev) =>
  dropzone.addEventListener(ev, (e) => {
    e.preventDefault();
    dropzone.classList.remove("dragover");
  })
);
dropzone.addEventListener("drop", (e) => {
  const file = e.dataTransfer.files[0];
  if (file) aceitarArquivo(file);
});

// ===================== PRÉVIA DA PLANILHA =====================
async function carregarPreview() {
  if (fileInput.files.length === 0) return;
  fileMessage.textContent = "Lendo planilha...";
  fileMessage.className = "field-hint";

  const body = new FormData();
  body.append("csrf_token", csrf());
  body.append("mode", modoAtual());
  body.append("planilha", fileInput.files[0]);

  try {
    const resp = await fetch("/preview", { method: "POST", body });
    const data = await resp.json();
    if (!resp.ok) throw new Error(data.error || "Não foi possível ler a planilha.");

    sumTotal.textContent = data.total;
    sumValid.textContent = data.validos;
    sumIgnored.textContent = data.ignorados;
    ignoradosPreview = data.ignorados;
    metricIgnored.textContent = data.ignorados;
    previewValido = data.validos > 0;
    fileMessage.textContent = `Planilha lida: ${data.total} linha(s), ${data.validos} válida(s).`;
    fileMessage.className = previewValido ? "field-hint ok" : "field-hint error";

    renderPreviewTable(data);
    refreshSubmitState();
  } catch (err) {
    previewValido = false;
    fileMessage.textContent = err.message;
    fileMessage.className = "field-hint error";
    previewBlock.classList.add("hidden");
    refreshSubmitState();
  }
}

function renderPreviewTable(data) {
  const cols = data.colunas || [];
  const rows = data.rows || [];
  if (cols.length === 0 || rows.length === 0) {
    previewBlock.classList.add("hidden");
    return;
  }
  let html = "<thead><tr>";
  for (const c of cols) html += `<th>${escapeHtml(c)}</th>`;
  html += "</tr></thead><tbody>";
  for (const row of rows) {
    html += "<tr>";
    for (const c of cols) html += `<td>${escapeHtml(row[c] || "")}</td>`;
    html += "</tr>";
  }
  html += "</tbody>";
  previewTable.innerHTML = html;

  // Mensagem da previa: o backend ja decide o texto e se e "ok" (verde) ou
  // "aviso" (amarelo), de acordo com o modo (consulta so precisa de email).
  const msg = data.mensagem_previa || "";
  const status = data.status_previa || "";
  previewWarning.classList.remove("hidden", "inline-ok", "inline-warning");
  if (msg) {
    const ok = status === "ok";
    previewWarning.textContent = `${ok ? "✅" : "⚠️"} ${msg}`;
    previewWarning.classList.add(ok ? "inline-ok" : "inline-warning");
  } else {
    previewWarning.classList.add("hidden");
  }
  previewBlock.classList.remove("hidden");
}

function escapeHtml(s) {
  return String(s).replace(/[&<>"']/g, (c) =>
    ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c])
  );
}

// ===================== PROGRESSO + TEMPO =====================
function setProgressBar(processed, total) {
  const pct = total > 0 ? Math.min(100, Math.round((processed / total) * 100)) : 0;
  progressBar.style.width = `${pct}%`;
  progressPercent.textContent = `${pct}%`;
  progressLabel.textContent =
    total > 0 ? `Processando ${Math.min(processed + (processed < total ? 1 : 0), total)} de ${total}` : "Aguardando início";
}

function mmss(segundos) {
  segundos = Math.max(0, Math.round(segundos));
  const m = Math.floor(segundos / 60);
  const s = segundos % 60;
  return `${m}:${String(s).padStart(2, "0")}`;
}

function atualizarTempo(data) {
  const progress = data.progress;
  if (!data.started_at) {
    progressTime.textContent = "";
    return;
  }
  const inicio = new Date(data.started_at).getTime();
  const fim = data.finished_at ? new Date(data.finished_at).getTime() : Date.now();
  const decorrido = (fim - inicio) / 1000;
  let txt = `⏱️ ${mmss(decorrido)} decorrido`;
  if (progress && progress.processados > 0 && progress.total > progress.processados && !data.finished_at) {
    const taxa = decorrido / progress.processados;
    const restante = taxa * (progress.total - progress.processados);
    txt += ` · ~${mmss(restante)} restante`;
  }
  progressTime.textContent = txt;
}

// ===================== CARDS =====================
function updateMetrics(data) {
  const progress = data.progress;
  if (progress) {
    metricProcessed.textContent = `${progress.processados}/${progress.total}`;
    metricSuccess.textContent = progress.sucessos;
    metricPending.textContent = progress.pendentes;
    metricErrors.textContent = progress.erros;
    metricIgnored.textContent = ignoradosPreview;
    setProgressBar(progress.processados, progress.total);
  }
}

// ===================== ÚLTIMAS AÇÕES (amigável) =====================
function nomeCurto(email) {
  return email && email.includes("@") ? email : email || "lead";
}

// Traduz uma linha técnica do log para uma frase amigável. Retorna null se for ruído.
function traduzirLinha(linha) {
  let m;
  // Consulta: [12/51] linha 12 email [✓|✗|!] telefone?
  m = linha.match(/^\[(\d+)\/(\d+)\]\s+linha\s+(\d+)\s+(\S+)\s+\[(✓|✗|!)\](?:\s+(\S+))?/);
  if (m) {
    const email = nomeCurto(m[4]);
    const marca = m[5];
    if (marca === "✓") return { cls: "ok", txt: `✅ Telefone encontrado — ${email}${m[6] ? ` (${m[6]})` : ""}` };
    if (marca === "✗") return { cls: "warn", txt: `➖ Sem telefone para ${email}` };
    return { cls: "err", txt: `⚠️ Não consegui consultar ${email}` };
  }
  // Cadastro
  m = linha.match(/^\[CADASTRADO\]\s+linha\s+(\d+)/i);
  if (m) return { cls: "ok", txt: `✅ Lead cadastrado no Sigavi (linha ${m[1]})` };
  m = linha.match(/^\[DUPLICADO\]\s+linha\s+(\d+)/i);
  if (m) return { cls: "warn", txt: `➖ Lead já existia (linha ${m[1]})` };
  m = linha.match(/^\[NAO CADASTRADO\]\s+linha\s+(\d+):?\s*(.*)/i);
  if (m) return { cls: "warn", txt: `🚫 Lead ignorado (linha ${m[1]})${m[2] ? `: ${m[2]}` : ""}` };
  m = linha.match(/^\[ERRO[^\]]*\]\s*(.*)/i);
  if (m) return { cls: "err", txt: `⚠️ Erro: ${m[1] || "ver detalhes técnicos"}` };
  // Informativos úteis
  if (/^Automacao iniciada/i.test(linha)) return { cls: "info", txt: "▶ Automação iniciada" };
  if (/Session de busca por email pronta/i.test(linha)) return { cls: "info", txt: "🔐 Conectado ao Sigavi" };
  if (/^Consulta paralela/i.test(linha)) return { cls: "info", txt: "🔍 Iniciando busca dos telefones" };
  if (/Consulta concluida/i.test(linha)) return { cls: "info", txt: "🏁 Busca finalizada" };
  if (/^Automacao concluida/i.test(linha)) return { cls: "info", txt: "🏁 Automação concluída" };
  if (/Resumo enviado|planilha\(s\) enviados pelo WhatsApp/i.test(linha)) return { cls: "ok", txt: "📲 Resumo enviado no WhatsApp" };
  return null;
}

function renderActions(logs) {
  const linhas = (logs || []).join("").split("\n").slice(-200);
  const acoes = [];
  for (const l of linhas) {
    const t = traduzirLinha(l.trim());
    if (t) acoes.push(t);
  }
  const recentes = acoes.slice(-40);
  if (recentes.length === 0) {
    actionsFeed.innerHTML = `<li class="action-empty">As ações aparecem aqui durante a execução.</li>`;
    return;
  }
  actionsFeed.innerHTML = recentes.map((a) => `<li class="${a.cls}">${escapeHtml(a.txt)}</li>`).join("");
  actionsFeed.scrollTop = actionsFeed.scrollHeight;
}

// ===================== LOG TÉCNICO =====================
function renderLogs(logs) {
  logOutput.textContent = logs.length ? logs.join("") : "Sem logs ainda.";
  logOutput.scrollTop = logOutput.scrollHeight;
}

// ===================== PAINEL FINAL =====================
function mostrarResultado(data) {
  const p = data.progress || {};
  const erros = p.erros || 0;
  const status = data.status;
  resultPanel.className = "result-panel";

  if (status === "failed") {
    resultPanel.classList.add("is-error");
    resultIcon.textContent = "❌";
    resultTitle.textContent = "Execução finalizada com erro";
    resultSubtitle.textContent = "Confira os detalhes técnicos abaixo.";
  } else if (erros > 0) {
    resultPanel.classList.add("has-errors");
    resultIcon.textContent = "⚠️";
    resultTitle.textContent = status === "stopped" ? "Execução parada (com alguns erros)" : "Execução concluída com alguns erros";
    resultSubtitle.textContent = `${erros} item(ns) não puderam ser processados.`;
  } else {
    resultIcon.textContent = "🎉";
    resultTitle.textContent = status === "stopped" ? "Execução parada — resultados salvos" : "Execução concluída com sucesso";
    resultSubtitle.textContent = "Tudo certo! Baixe o relatório abaixo.";
  }

  resultTotal.textContent = p.processados != null ? `${p.processados}/${p.total}` : "—";
  resultSuccess.textContent = p.sucessos ?? 0;
  resultPending.textContent = p.pendentes ?? 0;
  resultErrors.textContent = erros;

  renderDownloads(data);

  // Botao "Reprocessar so os erros": so faz sentido quando sobrou erro.
  reprocessMessage.textContent = "";
  reprocessMessage.className = "field-hint";
  if (erros > 0 && currentJobId) {
    reprocessJobId = currentJobId;
    reprocessButton.textContent = `🔁 Reprocessar ${erros} erro(s)`;
    reprocessButton.disabled = false;
    reprocessButton.classList.remove("hidden");
  } else {
    reprocessJobId = null;
    reprocessButton.classList.add("hidden");
  }

  resultPanel.classList.remove("hidden");
}

function renderDownloads(data) {
  const files = data.result_files || [];
  downloadList.innerHTML = "";
  if (!data.download_available || files.length === 0 || !currentJobId) return;
  for (const file of files) {
    const link = document.createElement("a");
    link.className = "download-button";
    link.href = `/jobs/${currentJobId}/download/${file.id}`;
    link.textContent = `⬇ ${file.filename}`;
    downloadList.appendChild(link);
  }
}

// ===================== POLLING =====================
async function pollJob() {
  if (!currentJobId) return;
  const response = await fetch(`/jobs/${currentJobId}`);
  const data = await response.json();

  if (!response.ok) {
    setStatus("Erro", "error");
    renderLogs([data.error || "Não foi possível buscar a execução."]);
    currentJobId = null;
    setBusy(false);
    return;
  }

  renderLogs(data.logs || []);
  renderActions(data.logs || []);
  updateMetrics(data);
  atualizarTempo(data);

  if (data.status === "queued") {
    setStatus("Na fila", "running");
  } else if (data.status === "running") {
    setStatus("Executando", "running");
  } else if (data.status === "stopping") {
    setStatus("Parando", "running");
  } else if (data.status === "completed" || data.status === "stopped") {
    setStatus(data.status === "stopped" ? "Parado" : "Finalizado", "done");
    mostrarResultado(data);
    currentJobId = null;
    setBusy(false);
    return;
  } else {
    setStatus("Erro", "error");
    mostrarResultado(data);
    currentJobId = null;
    setBusy(false);
    return;
  }
  pollTimer = window.setTimeout(pollJob, 1200);
}

// ===================== SUBMIT =====================
form.addEventListener("submit", async (event) => {
  event.preventDefault();
  window.clearTimeout(pollTimer);
  resultPanel.classList.add("hidden");

  setBusy(true);
  setStatus("Enviando", "running");
  jobIdLabel.textContent = "Preparando";
  logOutput.textContent = "Enviando planilha...\n";

  try {
    const response = await fetch("/jobs", { method: "POST", body: new FormData(form) });
    const data = await response.json();
    if (!response.ok) throw new Error(data.error || "Falha ao iniciar automação.");

    currentJobId = data.job_id;
    jobIdLabel.textContent = `#${currentJobId.slice(0, 8)}`;
    setStatus("Executando", "running");
    pollJob();
  } catch (error) {
    setStatus("Erro", "error");
    logOutput.textContent = `${error.message}\n`;
    currentJobId = null;
    setBusy(false);
  }
});

// ===================== PARAR =====================
stopButton.addEventListener("click", async () => {
  if (!currentJobId || stopButton.disabled) return;
  stopButton.disabled = true;
  setStatus("Parando", "running");

  const body = new FormData();
  body.append("csrf_token", csrf());
  try {
    const response = await fetch(`/jobs/${currentJobId}/stop`, { method: "POST", body });
    const data = await response.json();
    if (!response.ok) throw new Error(data.error || "Não foi possível parar a execução.");
  } catch (error) {
    setStatus("Erro", "error");
    logOutput.textContent += `\n${error.message}\n`;
  }
});

// ===================== REPROCESSAR ERROS =====================
reprocessButton.addEventListener("click", async () => {
  if (!reprocessJobId || reprocessButton.disabled) return;

  if (!loginInput.value.trim() || !senhaInput.value.trim()) {
    reprocessMessage.textContent = "Preencha login e senha do Sigavi para reprocessar.";
    reprocessMessage.className = "field-hint error";
    return;
  }

  reprocessButton.disabled = true;
  reprocessMessage.textContent = "Preparando reprocessamento...";
  reprocessMessage.className = "field-hint";

  const body = new FormData();
  body.append("csrf_token", csrf());
  body.append("sigavi_login", loginInput.value.trim());
  body.append("sigavi_senha", senhaInput.value);
  if (form.querySelector("[name='headless']").checked) body.append("headless", "on");

  try {
    const response = await fetch(`/jobs/${reprocessJobId}/reprocess`, { method: "POST", body });
    const data = await response.json();
    if (!response.ok) throw new Error(data.error || "Nao foi possivel reprocessar.");

    // Comeca a acompanhar o novo job, como num submit normal.
    window.clearTimeout(pollTimer);
    resultPanel.classList.add("hidden");
    reprocessJobId = null;
    setBusy(true);
    currentJobId = data.job_id;
    jobIdLabel.textContent = `#${currentJobId.slice(0, 8)}`;
    setStatus("Executando", "running");
    logOutput.textContent = `Reprocessando ${data.reprocessadas} linha(s) com erro...\n`;
    pollJob();
  } catch (error) {
    reprocessButton.disabled = false;
    reprocessMessage.textContent = error.message;
    reprocessMessage.className = "field-hint error";
  }
});

// ===================== LIMPAR TELA =====================
clearButton.addEventListener("click", () => {
  if (currentJobId) return; // não limpa durante execução
  window.clearTimeout(pollTimer);
  limparArquivo();
  setStatus("Pronto", "idle");
  sumStatus.textContent = "Aguardando";
  jobIdLabel.textContent = "Sem execução";
  logOutput.textContent = "Aguardando planilha.";
  setProgressBar(0, 0);
  progressTime.textContent = "";
  [metricProcessed, metricSuccess, metricPending, metricErrors, metricIgnored].forEach((el) => (el.textContent = "0"));
  metricProcessed.textContent = "0";
  resultPanel.classList.add("hidden");
  reprocessButton.classList.add("hidden");
  reprocessMessage.textContent = "";
  reprocessJobId = null;
  actionsFeed.innerHTML = `<li class="action-empty">As ações aparecem aqui durante a execução.</li>`;
});

// ===================== DESTINOS (enviar relatório para) =====================
async function carregarDestinos() {
  const container = document.querySelector("#destinos-list");
  if (!container) return;
  try {
    const resp = await fetch("/destinos");
    if (!resp.ok) throw new Error("falha");
    const data = await resp.json();
    const destinos = data.destinos || [];
    container.innerHTML = "";
    if (!destinos.length) {
      const p = document.createElement("p");
      p.className = "field-hint";
      p.textContent = "Nenhum destino configurado.";
      container.appendChild(p);
      return;
    }
    destinos.forEach((d) => {
      const ehGrupo = String(d.id).endsWith("@g.us");
      const label = document.createElement("label");
      label.className = "checkbox-row destino-item";
      const input = document.createElement("input");
      input.type = "checkbox";
      input.name = "destinos";
      input.value = d.id;
      if (d.padrao) input.checked = true;
      const span = document.createElement("span");
      span.textContent = `${ehGrupo ? "👥" : "📱"} ${d.nome}`;
      label.appendChild(input);
      label.appendChild(span);
      container.appendChild(label);
    });
  } catch (e) {
    container.innerHTML = "";
    const p = document.createElement("p");
    p.className = "field-hint";
    p.textContent = "Não foi possível carregar os destinos.";
    container.appendChild(p);
  }
}

// ===================== INICIALIZAÇÃO =====================
aplicarModo();
refreshSubmitState();
carregarDestinos();
