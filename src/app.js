const MAX_SIZE = 50 * 1024 * 1024; // 50 MB

const { invoke } = window.__TAURI__.core;
const { listen } = window.__TAURI__.event;

const els = {
  dropzone: document.getElementById("dropzone"),
  fileInput: document.getElementById("file-input"),
  file: document.getElementById("file"),
  fileName: document.getElementById("file-name"),
  fileSize: document.getElementById("file-size"),
  fileRemove: document.getElementById("file-remove"),
  convertBtn: document.getElementById("convert-btn"),
  stepSelect: document.getElementById("step-select"),
  stepWorking: document.getElementById("step-working"),
  stepDone: document.getElementById("step-done"),
  workingTitle: document.getElementById("working-title"),
  log: document.getElementById("log"),
  progressFill: document.getElementById("progress-fill"),
  doneSub: document.getElementById("done-sub"),
  downloadBtn: document.getElementById("download-btn"),
  againBtn: document.getElementById("again-btn"),
  error: document.getElementById("error"),
  errorText: document.getElementById("error-text"),
};

let selectedFile = null;
let resultBlob = null;
let resultName = "converted.docx";
let unlistenProgress = null;

const CHECK_SVG =
  '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3" stroke-linecap="round" stroke-linejoin="round"><polyline points="20 6 9 17 4 12"/></svg>';
let activeStep = null;

function formatSize(bytes) {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}

function showError(message) {
  els.errorText.textContent = message;
  els.error.hidden = false;
}
function clearError() {
  els.error.hidden = true;
}
function showStep(name) {
  els.stepSelect.hidden = name !== "select";
  els.stepWorking.hidden = name !== "working";
  els.stepDone.hidden = name !== "done";
}

function clearLog() {
  els.log.innerHTML = "";
  activeStep = null;
  setProgress(0);
}
function setProgress(p) {
  if (p == null) return;
  els.progressFill.style.width = `${Math.max(0, Math.min(100, p))}%`;
}

async function addStep(step, detail) {
  if (activeStep === step) {
    const last = els.log.lastElementChild;
    if (last && detail) {
      const d = last.querySelector(".detail");
      if (d) d.textContent = detail;
    }
    return;
  }
  const prev = els.log.lastElementChild;
  if (prev) {
    prev.classList.remove("active");
    prev.classList.add("done");
    prev.querySelector(".icon").innerHTML = CHECK_SVG;
  }
  activeStep = step;
  const li = document.createElement("li");
  li.className = "active";
  const icon = document.createElement("span");
  icon.className = "icon";
  icon.innerHTML = '<span class="spin"></span>';
  const label = document.createElement("span");
  label.className = "label";
  label.textContent = step;
  li.appendChild(icon);
  li.appendChild(label);
  if (detail) {
    const d = document.createElement("span");
    d.className = "detail";
    d.textContent = detail;
    li.appendChild(d);
  }
  els.log.appendChild(li);
  await new Promise((r) => setTimeout(r, 40));
}

function markAllDone() {
  const prev = els.log.lastElementChild;
  if (prev) {
    prev.classList.remove("active");
    prev.classList.add("done");
    prev.querySelector(".icon").innerHTML = CHECK_SVG;
  }
  activeStep = null;
}

function setFile(file) {
  if (!file) return;
  if (file.type !== "application/pdf" && !file.name.toLowerCase().endsWith(".pdf")) {
    showError("That doesn't look like a PDF. Please choose a .pdf file.");
    return;
  }
  if (file.size > MAX_SIZE) {
    showError("File is too large. The limit is 50 MB.");
    return;
  }
  clearError();
  selectedFile = file;
  els.fileName.textContent = file.name;
  els.fileSize.textContent = formatSize(file.size);
  els.file.hidden = false;
  els.convertBtn.disabled = false;
}

function reset() {
  selectedFile = null;
  resultBlob = null;
  els.fileInput.value = "";
  els.file.hidden = true;
  els.convertBtn.disabled = true;
  clearError();
  showStep("select");
}

// --- Wiring ---

els.dropzone.addEventListener("click", () => els.fileInput.click());
els.dropzone.addEventListener("keydown", (e) => {
  if (e.key === "Enter" || e.key === " ") {
    e.preventDefault();
    els.fileInput.click();
  }
});
els.fileInput.addEventListener("change", () => {
  if (els.fileInput.files?.length) setFile(els.fileInput.files[0]);
});
els.fileRemove.addEventListener("click", () => {
  selectedFile = null;
  els.fileInput.value = "";
  els.file.hidden = true;
  els.convertBtn.disabled = true;
});
["dragenter", "dragover"].forEach((ev) =>
  els.dropzone.addEventListener(ev, (e) => {
    e.preventDefault();
    els.dropzone.classList.add("dragging");
  }),
);
["dragleave", "drop"].forEach((ev) =>
  els.dropzone.addEventListener(ev, (e) => {
    e.preventDefault();
    els.dropzone.classList.remove("dragging");
  }),
);
els.dropzone.addEventListener("drop", (e) => {
  const files = e.dataTransfer?.files;
  if (files?.length) setFile(files[0]);
});
els.convertBtn.addEventListener("click", convert);
els.againBtn.addEventListener("click", reset);
els.downloadBtn.addEventListener("click", download);

async function convert() {
  if (!selectedFile) return;

  clearError();
  clearLog();
  showStep("working");
  addStep("读取文件", formatSize(selectedFile.size));

  try {
    const bytes = new Uint8Array(await selectedFile.arrayBuffer());

    // Listen for progress events emitted by the Rust backend.
    unlistenProgress = await listen("progress", (e) => {
      const payload = e.payload;
      addStep(payload.step, payload.detail || "");
      const m = /第 (\d+) \/ (\d+) 页/.exec(payload.detail || "");
      if (m) setProgress(Math.round((Number(m[1]) / Number(m[2])) * 90));
    });

    const result = await invoke("convert_pdf", { data: bytes });
    unlistenProgress();
    unlistenProgress = null;

    const bytesOut = result instanceof Uint8Array ? result : new Uint8Array(result);
    resultBlob = new Blob([bytesOut], {
      type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    });
    resultName = selectedFile.name.replace(/\.pdf$/i, "") + ".docx";

    markAllDone();
    setProgress(100);
    els.doneSub.textContent = `${formatSize(resultBlob.size)} · ready to save`;
    showStep("done");
  } catch (err) {
    if (unlistenProgress) {
      unlistenProgress();
      unlistenProgress = null;
    }
    showStep("select");
    showError(err?.message || String(err) || "Conversion failed.");
  }
}

function download() {
  if (!resultBlob) return;
  const url = URL.createObjectURL(resultBlob);
  const a = document.createElement("a");
  a.href = url;
  a.download = resultName;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 4000);
}

window.addEventListener("dragover", (e) => e.preventDefault());
window.addEventListener("drop", (e) => {
  if (e.target.closest("#dropzone")) return;
  const files = e.dataTransfer?.files;
  if (files?.length) {
    e.preventDefault();
    setFile(files[0]);
  }
});
