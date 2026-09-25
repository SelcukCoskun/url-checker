"use strict";

const $ = (selector) => document.querySelector(selector);
const token = document.querySelector('meta[name="app-token"]').content;
const model = { datasetId: "", filename: "", urlCount: 0, jobId: "", results: [], total: 0, status: "idle", toastTimer: 0 };

const fileInput = $("#file-input");
const dropzone = $("#dropzone");
const fileCard = $("#file-card");
const removeFileButton = $("#remove-file");
const scanButton = $("#scan-button");
const exportButton = $("#export-button");

async function localRequest(path, options = {}) {
  const headers = new Headers(options.headers || {});
  headers.set("X-App-Token", token);
  const response = await fetch(path, { ...options, headers, cache: "no-store", credentials: "same-origin" });
  if (!response.ok) {
    let message = `İstek başarısız oldu (${response.status}).`;
    try { message = (await response.json()).error || message; } catch (_) { /* keep status text */ }
    throw new Error(message);
  }
  return response;
}

function toast(message, success = false) {
  const node = $("#toast");
  node.textContent = message;
  node.classList.toggle("is-success", success);
  node.classList.add("is-visible");
  window.clearTimeout(model.toastTimer);
  model.toastTimer = window.setTimeout(() => node.classList.remove("is-visible"), 3900);
}

function setUploadMessage(message, isError = false) {
  const node = $("#upload-message");
  node.textContent = message;
  node.style.color = isError ? "#b7474e" : "";
}

function setScanButton(label, disabled) {
  scanButton.querySelector("span").textContent = label;
  scanButton.disabled = disabled;
}

function resetResults() {
  model.jobId = "";
  model.results = [];
  model.total = model.urlCount;
  model.status = "ready";
  $("#results-body").replaceChildren();
  $("#result-counter").textContent = "Henüz tarama yok";
  $("#empty-state").classList.remove("hidden");
  $("#empty-state strong").textContent = "Sonuçlar burada görünecek";
  $("#empty-state > span:last-child").textContent = "Bir dosya seç, adresleri hazırla ve kontrolü başlat.";
  $("#table-footer").classList.add("hidden");
  $("#progress-area").classList.add("hidden");
  exportButton.disabled = true;
  updateStats();
  renderFilters();
}

async function clearDataset() {
  try { await localRequest("/api/reset", { method: "POST", headers: { "Content-Type": "application/json" }, body: "{}" }); }
  catch (_) { /* server may already have stopped */ }
  model.datasetId = "";
  model.filename = "";
  model.urlCount = 0;
  model.total = 0;
  model.status = "idle";
  fileCard.classList.add("hidden");
  dropzone.classList.remove("hidden");
  $("#preview-list").classList.add("hidden");
  $("#preview-list").replaceChildren();
  fileInput.value = "";
  setScanButton("Kontrolü başlat", true);
  setUploadMessage("Kontrol, sen başlatana kadar başlamaz.");
  resetResults();
  updateStats();
}

function showFileCard(data, file) {
  model.datasetId = data.dataset_id;
  model.filename = data.filename;
  model.urlCount = data.count;
  model.total = data.count;
  $("#file-name").textContent = data.filename;
  const extension = file.name.slice(file.name.lastIndexOf(".")).toLowerCase();
  const readMode = extension === ".txt" ? "satırlardan adresler okundu" : "ilk sütundaki adresler okundu";
  $("#file-caption").textContent = `${(file.size / 1024).toFixed(1)} KB · ${readMode}`;
  $("#file-count").textContent = `${data.count} adres`;
  fileCard.classList.remove("hidden");
  removeFileButton.disabled = false;
  dropzone.classList.add("hidden");
  const preview = $("#preview-list");
  preview.replaceChildren();
  for (const address of data.preview || []) {
    const chip = document.createElement("span");
    chip.className = "preview-chip";
    chip.textContent = address;
    chip.title = address;
    preview.append(chip);
  }
  if (data.count > (data.preview || []).length) {
    const more = document.createElement("span");
    more.className = "preview-more";
    more.textContent = `+${data.count - data.preview.length} adres daha`;
    preview.append(more);
  }
  preview.classList.toggle("hidden", !data.count);
  setScanButton("Kontrolü başlat", false);
  setUploadMessage("Dosya hazır. Kontrol yalnızca başlat düğmesine bastığında yapılır.");
  resetResults();
  updateStats();
}

async function uploadFile(file) {
  if (!file) return;
  const supportedExtensions = [".xlsx", ".xlsm", ".xls", ".ods", ".csv", ".tsv", ".txt"];
  const extension = file.name.slice(file.name.lastIndexOf(".")).toLowerCase();
  if (!supportedExtensions.includes(extension)) {
    toast("XLSX, XLSM, XLS, ODS, CSV, TSV veya TXT biçiminde dosya seç.");
    return;
  }
  if (file.size > 10 * 1024 * 1024) {
    toast("Dosya boyutu 10 MB sınırını aşıyor.");
    return;
  }
  model.datasetId = "";
  model.filename = file.name;
  model.urlCount = 0;
  model.total = 0;
  model.status = "uploading";
  resetResults();
  dropzone.classList.add("hidden");
  fileCard.classList.remove("hidden");
  removeFileButton.disabled = true;
  $("#file-name").textContent = file.name;
  $("#file-caption").textContent = "Dosya yerel olarak okunuyor…";
  $("#file-count").textContent = "Hazırlanıyor";
  setScanButton("Dosya okunuyor…", true);
  setUploadMessage("Dosya bu bilgisayarda inceleniyor.");
  try {
    const response = await localRequest("/api/upload", {
      method: "POST",
      headers: { "Content-Type": "application/octet-stream", "X-File-Name": encodeURIComponent(file.name) },
      body: file,
    });
    const data = await response.json();
    showFileCard(data, file);
    toast(`${data.count} adres hazır. Tarama henüz başlamadı.`, true);
  } catch (error) {
    fileCard.classList.add("hidden");
    dropzone.classList.remove("hidden");
    removeFileButton.disabled = false;
    fileInput.value = "";
    model.status = "idle";
    setScanButton("Kontrolü başlat", true);
    setUploadMessage("Dosya okunamadı.", true);
    toast(error.message || "Dosya okunamadı.");
  }
}

function stateLabel(result) {
  if (result.availability) return result.availability;
  if (result.code === "403") return "Yanıt veriyor · erişim kısıtlı";
  if (result.code === "401") return "Aktif · oturum gerekli";
  if (result.code === "404" || result.code === "410") return "Yanıt veriyor · sayfa bulunamadı";
  const labels = { Aktif: "Aktif", Ulaşılamıyor: "Ulaşılamıyor", Belirsiz: "Belirsiz", Engellendi: "Güvenlik nedeniyle engellendi", Geçersiz: "Geçersiz adres" };
  return labels[result.state] || result.state || "Belirsiz";
}

function stateClass(result) {
  const label = stateLabel(result);
  if (label.startsWith("Aktif")) return "state-active";
  if (/sunucu hatası/i.test(label)) return "state-error";
  if (label.startsWith("Yanıt veriyor")) return "state-restricted";
  if (result.state === "Engellendi") return "state-unknown";
  if (result.state === "Ulaşılamıyor" || result.state === "Geçersiz") return "state-error";
  return "state-unknown";
}

function responseCodes(result) {
  const chain = Array.isArray(result.response_chain) ? result.response_chain : [];
  return [...new Set([...chain, result.code].filter(Boolean).map(String))];
}

function makeCell(className, text, title = "") {
  const cell = document.createElement("td");
  if (className) cell.className = className;
  cell.textContent = text || "—";
  if (title) cell.title = title;
  return cell;
}

function safeWebAddress(value) {
  try {
    const parsed = new URL(value);
    return parsed.protocol === "https:" || parsed.protocol === "http:" ? parsed.href : "";
  } catch (_) { return ""; }
}

function renderRow(result) {
  const row = document.createElement("tr");
  const urlCell = document.createElement("td");
  urlCell.className = "url-cell";
  const safeAddress = safeWebAddress(result.final_url || result.url || "");
  if (safeAddress) {
    const link = document.createElement("a");
    link.className = "url-main";
    link.href = safeAddress;
    link.target = "_blank";
    link.rel = "noopener noreferrer";
    link.textContent = result.url || safeAddress;
    link.title = "Yeni sekmede aç";
    urlCell.append(link);
  } else {
    const label = document.createElement("span");
    label.className = "url-main";
    label.textContent = result.url || "—";
    label.title = label.textContent;
    urlCell.append(label);
  }
  if (result.final_url && result.final_url !== result.url) {
    const final = document.createElement("span");
    final.className = "url-sub";
    final.textContent = `Son adres: ${result.final_url}`;
    final.title = final.textContent;
    urlCell.append(final);
  }
  row.append(urlCell);

  const codeCell = document.createElement("td");
  const flow = document.createElement("span");
  flow.className = "response-flow";
  const codes = responseCodes(result);
  codes.forEach((value, index) => {
    if (index) flow.append(document.createTextNode(" → "));
    const code = document.createElement("span");
    code.className = "code-pill";
    if (/^2/.test(value)) code.classList.add("code-2xx");
    else if (/^3/.test(value)) code.classList.add("code-3xx");
    else if (/^4/.test(value)) code.classList.add("code-4xx");
    else if (/^5/.test(value)) code.classList.add("code-5xx");
    else code.classList.add("code-none");
    code.textContent = value || "Kod yok";
    flow.append(code);
  });
  codeCell.append(flow);
  const hops = Array.isArray(result.redirect_chain) ? result.redirect_chain : [];
  codeCell.title = hops.map((hop) => `${hop.code}: ${hop.from} → ${hop.to}`).join("\n") || `Son HTTP yanıtı: ${result.code || "kod yok"}`;
  row.append(codeCell);

  const statusCell = document.createElement("td");
  const status = document.createElement("span");
  status.className = `state-badge ${stateClass(result)}`;
  const dot = document.createElement("span");
  dot.className = "state-dot";
  status.append(dot, document.createTextNode(stateLabel(result)));
  statusCell.append(status);
  row.append(statusCell);
  row.append(makeCell("tls-label", result.tls, result.tls));
  row.append(makeCell("title-cell", result.title, result.title));
  row.append(makeCell("summary-cell", result.summary || result.detail, result.summary || result.detail));

  const openCell = document.createElement("td");
  const openUrl = safeWebAddress(result.final_url || result.url || "");
  if (openUrl) {
    const openLink = document.createElement("a");
    openLink.className = "open-link";
    openLink.href = openUrl;
    openLink.target = "_blank";
    openLink.rel = "noopener noreferrer";
    openLink.title = "Siteyi yeni sekmede aç";
    openLink.setAttribute("aria-label", "Siteyi yeni sekmede aç");
    openLink.innerHTML = '<svg viewBox="0 0 24 24" fill="none" aria-hidden="true"><path d="M13.5 5H19v5.5M18.5 5.5l-8 8M18 13v4.5a1.5 1.5 0 0 1-1.5 1.5h-10A1.5 1.5 0 0 1 5 17.5v-10A1.5 1.5 0 0 1 6.5 6H11" stroke="currentColor" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round"/></svg>';
    openCell.append(openLink);
  }
  row.append(openCell);
  row.dataset.search = [result.url, result.final_url, result.title, result.summary, result.detail, ...codes, stateLabel(result)].join(" ").toLocaleLowerCase("tr-TR");
  row.dataset.state = stateLabel(result);
  row.dataset.codes = codes.join("|");
  return row;
}

function renderFilters() {
  const stateFilter = $("#state-filter");
  const codeFilter = $("#code-filter");
  const priorState = stateFilter.value;
  const priorCode = codeFilter.value;
  const states = [...new Set(model.results.map(stateLabel).filter(Boolean))];
  const codes = [...new Set(model.results.flatMap(responseCodes).filter(Boolean))].sort((a, b) => {
    const aa = Number(a), bb = Number(b);
    return Number.isNaN(aa) || Number.isNaN(bb) ? a.localeCompare(b, "tr") : aa - bb;
  });
  stateFilter.replaceChildren(new Option("Tüm durumlar", ""));
  for (const value of states) stateFilter.add(new Option(value, value));
  codeFilter.replaceChildren(new Option("Tüm kodlar", ""));
  for (const value of codes) codeFilter.add(new Option(value, value));
  if (states.includes(priorState)) stateFilter.value = priorState;
  if (codes.includes(priorCode)) codeFilter.value = priorCode;
}

function updateStats() {
  $("#stat-total").textContent = model.total ? String(model.total) : "—";
  const complete = model.results.length;
  $("#stat-active").textContent = complete ? String(model.results.filter((row) => row.state === "Aktif").length) : "—";
  $("#stat-403").textContent = complete ? String(model.results.filter((row) => row.code === "403").length) : "—";
  $("#stat-pending").textContent = model.total ? String(Math.max(0, model.total - complete)) : "—";
}

function applyFilters() {
  const query = $("#search-input").value.trim().toLocaleLowerCase("tr-TR");
  const state = $("#state-filter").value;
  const code = $("#code-filter").value;
  const body = $("#results-body");
  body.replaceChildren();
  let shown = 0;
  for (const result of model.results) {
    if (state && stateLabel(result) !== state) continue;
    if (code && !responseCodes(result).includes(code)) continue;
    const row = renderRow(result);
    if (query && !row.dataset.search.includes(query)) continue;
    body.append(row);
    shown += 1;
  }
  const hasResults = model.results.length > 0;
  const noneVisible = hasResults && shown === 0;
  $("#empty-state").classList.toggle("hidden", hasResults && !noneVisible);
  if (noneVisible) {
    $("#empty-state strong").textContent = "Filtreye uyan sonuç yok";
    $("#empty-state > span:last-child").textContent = "Arama metnini veya filtreleri değiştir.";
  }
  $("#table-footer").classList.toggle("hidden", !hasResults);
  $("#visible-count").textContent = `${shown} / ${model.results.length} sonuç gösteriliyor`;
  $("#result-counter").textContent = model.status === "running" ? `${model.results.length} / ${model.total} tamamlandı` : `${model.results.length} sonuç`;
}

function updateProgress(job) {
  const pct = job.total ? Math.round((job.completed / job.total) * 100) : 0;
  $("#progress-area").classList.remove("hidden");
  $("#progress-label").textContent = job.status === "completed" ? "Tarama tamamlandı" : "Adresler kontrol ediliyor";
  $("#progress-number").textContent = `${pct}% · ${job.completed}/${job.total}`;
  $("#progress-bar").style.width = `${pct}%`;
  model.total = job.total;
  model.results = job.results || [];
  model.status = job.status;
  renderFilters();
  applyFilters();
  updateStats();
  exportButton.disabled = model.results.length === 0;
  if (job.status === "running" || job.status === "queued") {
    removeFileButton.disabled = true;
    setScanButton("Tarama sürüyor…", true);
  } else if (job.status === "completed") {
    removeFileButton.disabled = false;
    setScanButton("Yeniden tara", false);
    setUploadMessage("Tarama tamamlandı. Sonuçları filtreleyebilir veya XLSX olarak indirebilirsin.");
  } else if (job.status === "error") {
    removeFileButton.disabled = false;
    setScanButton("Tekrar dene", false);
    setUploadMessage(job.error || "Tarama tamamlanamadı.", true);
  }
}

async function pollJob(jobId) {
  if (model.jobId !== jobId) return;
  try {
    const response = await localRequest(`/api/jobs/${encodeURIComponent(jobId)}`);
    const job = await response.json();
    updateProgress(job);
    if (job.status === "running" || job.status === "queued") {
      window.setTimeout(() => pollJob(jobId), 850);
    } else if (job.status === "completed") {
      toast(`${job.completed} adres için kontrol tamamlandı.`, true);
    } else {
      toast(job.error || "Tarama tamamlanamadı.");
    }
  } catch (error) {
    model.status = "error";
    setScanButton("Tekrar dene", false);
    setUploadMessage(error.message, true);
    toast(error.message);
  }
}

async function startScan() {
  if (!model.datasetId || scanButton.disabled) return;
  setScanButton("Tarama başlatılıyor…", true);
  setUploadMessage("Yalnızca seçilen dosyadaki herkese açık web adresleri kontrol edilecek.");
  try {
    const response = await localRequest("/api/scan", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ dataset_id: model.datasetId }),
    });
    const data = await response.json();
    model.jobId = data.job_id;
    model.total = data.total;
    model.results = [];
    model.status = "running";
    removeFileButton.disabled = true;
    $("#progress-bar").style.width = "0%";
    $("#progress-area").classList.remove("hidden");
    exportButton.disabled = true;
    updateStats();
    pollJob(model.jobId);
  } catch (error) {
    setScanButton("Kontrolü başlat", false);
    setUploadMessage(error.message, true);
    toast(error.message);
  }
}

async function exportResults() {
  if (!model.jobId || !model.results.length) return;
  try {
    const response = await localRequest(`/api/jobs/${encodeURIComponent(model.jobId)}/export`);
    const blob = await response.blob();
    const objectUrl = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = objectUrl;
    link.download = "domain-kontrol-sonuclari.xlsx";
    document.body.append(link);
    link.click();
    link.remove();
    window.setTimeout(() => URL.revokeObjectURL(objectUrl), 1000);
    toast("XLSX raporu indirildi.", true);
  } catch (error) { toast(error.message); }
}

$("#browse-button").addEventListener("click", () => fileInput.click());
dropzone.addEventListener("click", (event) => {
  if (!event.target.closest("button")) fileInput.click();
});
dropzone.addEventListener("keydown", (event) => {
  if (event.key === "Enter" || event.key === " ") { event.preventDefault(); fileInput.click(); }
});
fileInput.addEventListener("change", () => uploadFile(fileInput.files && fileInput.files[0]));
$("#remove-file").addEventListener("click", clearDataset);
scanButton.addEventListener("click", startScan);
exportButton.addEventListener("click", exportResults);
$("#search-input").addEventListener("input", applyFilters);
$("#state-filter").addEventListener("change", applyFilters);
$("#code-filter").addEventListener("change", applyFilters);
$("#clear-filters").addEventListener("click", () => {
  $("#search-input").value = "";
  $("#state-filter").value = "";
  $("#code-filter").value = "";
  applyFilters();
});

for (const eventName of ["dragenter", "dragover"]) {
  dropzone.addEventListener(eventName, (event) => { event.preventDefault(); dropzone.classList.add("is-dragging"); });
}
for (const eventName of ["dragleave", "dragend"]) {
  dropzone.addEventListener(eventName, () => dropzone.classList.remove("is-dragging"));
}
dropzone.addEventListener("drop", (event) => {
  event.preventDefault();
  dropzone.classList.remove("is-dragging");
  uploadFile(event.dataTransfer && event.dataTransfer.files && event.dataTransfer.files[0]);
});

updateStats();
