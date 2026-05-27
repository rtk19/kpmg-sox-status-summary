const LOADING_STEPS = [
  "מנתח מבנה הגיליונות...",
  "מזהה עמודות סטטוס ומפתח...",
  "מחשב סיכומים...",
  "בונה גיליונות פלט...",
  "שומר קובץ...",
];

let files = [];
let view = "upload";
let progress = 0;
let stepIdx = 0;
let tick = null;
let download = null;

const uploadView = document.getElementById("uploadView");
const loadingView = document.getElementById("loadingView");
const successView = document.getElementById("successView");
const dropzone = document.getElementById("dropzone");
const fileInput = document.getElementById("fileInput");
const browseButton = document.getElementById("browseButton");
const fileList = document.getElementById("fileList");
const submitButton = document.getElementById("submitButton");
const errorBox = document.getElementById("errorBox");
const errorText = document.getElementById("errorText");
const loadingStep = document.getElementById("loadingStep");
const progressBar = document.getElementById("progressBar");
const redownloadButton = document.getElementById("redownloadButton");
const resetButton = document.getElementById("resetButton");

const helpDialog = document.getElementById("helpDialog");
const helpButton = document.getElementById("helpButton");
const closeHelpButton = document.getElementById("closeHelpButton");
const howTab = document.getElementById("howTab");
const templateTab = document.getElementById("templateTab");
const howContent = document.getElementById("howContent");
const templateContent = document.getElementById("templateContent");

function formatBytes(bytes) {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}

function setView(nextView) {
  view = nextView;
  uploadView.hidden = nextView !== "upload";
  loadingView.hidden = nextView !== "loading";
  successView.hidden = nextView !== "success";
}

function setError(message) {
  errorText.textContent = message || "";
  errorBox.hidden = !message;
}

function renderFiles() {
  fileList.innerHTML = "";
  files.forEach((file, index) => {
    const item = document.createElement("li");
    item.className = "file-item";
    item.innerHTML = `
      <svg viewBox="0 0 24 24" aria-hidden="true"><path d="M15 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V7Z"></path><path d="M14 2v4a2 2 0 0 0 2 2h4"></path><path d="M8 13h8"></path><path d="M8 17h8"></path><path d="M10 9H8"></path></svg>
      <div class="file-meta">
        <div class="file-name"></div>
        <div class="file-size"></div>
      </div>
      <button type="button" class="remove-file" aria-label="הסר קובץ">
        <svg viewBox="0 0 24 24" aria-hidden="true"><path d="M18 6 6 18"></path><path d="m6 6 12 12"></path></svg>
      </button>
    `;
    item.querySelector(".file-name").textContent = file.name;
    item.querySelector(".file-size").textContent = formatBytes(file.size);
    item.querySelector(".remove-file").addEventListener("click", () => {
      files = files.filter((_, i) => i !== index);
      renderFiles();
    });
    fileList.appendChild(item);
  });

  submitButton.disabled = files.length === 0;
  submitButton.querySelector("span").textContent =
    files.length > 1 ? `שליחת ${files.length} קבצים לעיבוד` : "שליחה לעיבוד";
}

function addFiles(incoming) {
  const arr = Array.from(incoming);
  const valid = [];

  for (const file of arr) {
    const ext = file.name.split(".").pop()?.toLowerCase();
    if (ext !== "xlsx") {
      setError("יש להעלות קבצים מסוג xlsx בלבד");
      return;
    }
    valid.push(file);
  }

  setError("");
  const seen = new Set(files.map((file) => `${file.name}_${file.size}`));
  for (const file of valid) {
    const key = `${file.name}_${file.size}`;
    if (!seen.has(key)) {
      seen.add(key);
      files.push(file);
    }
  }
  renderFiles();
}

function clearFiles() {
  files = [];
  setError("");
  fileInput.value = "";
  renderFiles();
}

function startTicker() {
  progress = 0;
  stepIdx = 0;
  progressBar.style.width = "0%";
  loadingStep.textContent = LOADING_STEPS[0];
  tick = window.setInterval(() => {
    progress = Math.min(progress + Math.random() * 14 + 5, 92);
    stepIdx = Math.min(stepIdx + 1, LOADING_STEPS.length - 1);
    progressBar.style.width = `${Math.round(progress)}%`;
    loadingStep.textContent = LOADING_STEPS[stepIdx];
  }, 600);
}

function stopTicker() {
  if (tick) window.clearInterval(tick);
  tick = null;
}

function triggerDownload(url, name) {
  const link = document.createElement("a");
  link.href = url;
  link.download = name;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
}

function parseFilename(response) {
  const disposition = response.headers.get("Content-Disposition") || "";
  const utf8 = disposition.match(/filename\*=UTF-8''([^;]+)/i);
  if (utf8) return decodeURIComponent(utf8[1]);
  const ascii = disposition.match(/filename="?([^";]+)"?/i);
  if (ascii) return ascii[1];
  return files.length === 1
    ? `${files[0].name.replace(/\.xlsx$/i, "").replace(/\s+/g, "_")}_with_status.xlsx`
    : "sox_summary_of_controls_combined_with_status.xlsx";
}

async function handleSubmit() {
  if (!files.length) return;
  setView("loading");
  startTicker();

  const body = new FormData();
  files.forEach((file) => body.append("files", file, file.name));

  try {
    await new Promise((resolve) => window.setTimeout(resolve, 50));
    const response = await fetch("/process", { method: "POST", body });
    if (!response.ok) {
      let message = "שגיאה לא צפויה בעיבוד הקובץ.";
      try {
        const payload = await response.json();
        if (payload?.error) message = payload.error;
      } catch {}
      throw new Error(message);
    }

    const blob = await response.blob();
    const filename = parseFilename(response);
    stopTicker();
    progressBar.style.width = "100%";

    if (download?.url) URL.revokeObjectURL(download.url);
    const url = URL.createObjectURL(blob);
    download = { url, name: filename };
    triggerDownload(url, filename);
    window.setTimeout(() => setView("success"), 400);
  } catch (error) {
    stopTicker();
    setError(error.message || String(error));
    setView("upload");
  }
}

function reset() {
  setView("upload");
  clearFiles();
  if (download?.url) URL.revokeObjectURL(download.url);
  download = null;
}

dropzone.addEventListener("click", (event) => {
  if (event.target.tagName !== "BUTTON") fileInput.click();
});

dropzone.addEventListener("keydown", (event) => {
  if (event.key === "Enter" || event.key === " ") fileInput.click();
});

browseButton.addEventListener("click", (event) => {
  event.stopPropagation();
  fileInput.click();
});

fileInput.addEventListener("change", () => {
  if (fileInput.files?.length) addFiles(fileInput.files);
});

dropzone.addEventListener("dragover", (event) => {
  event.preventDefault();
  dropzone.classList.add("dragging");
});

dropzone.addEventListener("dragleave", () => {
  dropzone.classList.remove("dragging");
});

dropzone.addEventListener("drop", (event) => {
  event.preventDefault();
  dropzone.classList.remove("dragging");
  if (event.dataTransfer.files.length) addFiles(event.dataTransfer.files);
});

window.addEventListener("dragover", (event) => {
  if (!Array.from(event.dataTransfer?.types || []).includes("Files")) return;
  if (view === "loading") return;
  event.preventDefault();
  dropzone.classList.add("dragging");
});

window.addEventListener("dragleave", (event) => {
  if (event.relatedTarget === null) dropzone.classList.remove("dragging");
});

window.addEventListener("drop", (event) => {
  if (!event.dataTransfer?.files?.length || view === "loading") return;
  event.preventDefault();
  dropzone.classList.remove("dragging");
  if (view === "success") {
    if (download?.url) URL.revokeObjectURL(download.url);
    download = null;
    files = [];
    setView("upload");
  }
  addFiles(event.dataTransfer.files);
});

submitButton.addEventListener("click", handleSubmit);
resetButton.addEventListener("click", reset);
redownloadButton.addEventListener("click", () => {
  if (download) triggerDownload(download.url, download.name);
});

helpButton.addEventListener("click", () => {
  helpDialog.hidden = false;
});

closeHelpButton.addEventListener("click", () => {
  helpDialog.hidden = true;
});

helpDialog.addEventListener("click", (event) => {
  if (event.target === helpDialog) helpDialog.hidden = true;
});

function setTab(which) {
  const isHow = which === "how";
  howTab.classList.toggle("active", isHow);
  templateTab.classList.toggle("active", !isHow);
  howTab.setAttribute("aria-selected", String(isHow));
  templateTab.setAttribute("aria-selected", String(!isHow));
  howContent.hidden = !isHow;
  templateContent.hidden = isHow;
}

howTab.addEventListener("click", () => setTab("how"));
templateTab.addEventListener("click", () => setTab("template"));

window.addEventListener("beforeunload", () => {
  if (download?.url) URL.revokeObjectURL(download.url);
});

renderFiles();
