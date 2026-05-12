const form = document.querySelector("#uploadForm");
const fileInput = document.querySelector("#fileInput");
const fileName = document.querySelector("#fileName");
const fileMeta = document.querySelector("#fileMeta");
const statusText = document.querySelector("#status");
const submitButton = document.querySelector("#submitButton");
const result = document.querySelector("#result");
const downloadLink = document.querySelector("#downloadLink");
const reportList = document.querySelector("#reportList");
const warningList = document.querySelector("#warningList");
const dropzone = document.querySelector(".dropzone");
const fontSelects = document.querySelectorAll(".format-options select[data-default]:not(.size-select)");
const sizeSelects = document.querySelectorAll(".size-select");

const fallbackFonts = [
  "Times New Roman",
  "Arial",
  "Calibri",
  "宋体",
  "仿宋",
  "黑体",
  "楷体",
  "微软雅黑",
  "方正公文小标宋",
  "方正公文仿宋",
  "方正公文黑体",
  "方正公文楷体",
];

const chineseFontSizes = [
  ["42", "初号（42pt）"],
  ["36", "小初（36pt）"],
  ["26", "一号（26pt）"],
  ["24", "小一（24pt）"],
  ["22", "二号（22pt）"],
  ["18", "小二（18pt）"],
  ["16", "三号（16pt）"],
  ["15", "小三（15pt）"],
  ["14", "四号（14pt）"],
  ["12", "小四（12pt）"],
  ["10.5", "五号（10.5pt）"],
  ["9", "小五（9pt）"],
  ["7.5", "六号（7.5pt）"],
];

initializeOptions();

fileInput.addEventListener("change", () => {
  updateSelectedFile(fileInput.files[0]);
});

["dragenter", "dragover"].forEach((eventName) => {
  dropzone.addEventListener(eventName, (event) => {
    event.preventDefault();
    dropzone.classList.add("dragging");
  });
});

["dragleave", "drop"].forEach((eventName) => {
  dropzone.addEventListener(eventName, (event) => {
    event.preventDefault();
    dropzone.classList.remove("dragging");
  });
});

dropzone.addEventListener("drop", (event) => {
  const file = event.dataTransfer.files[0];
  if (!file) return;
  const transfer = new DataTransfer();
  transfer.items.add(file);
  fileInput.files = transfer.files;
  updateSelectedFile(file);
});

form.addEventListener("submit", async (event) => {
  event.preventDefault();
  statusText.textContent = "";
  result.hidden = true;

  const file = fileInput.files[0];
  if (!file) {
    statusText.textContent = "请先选择一个文稿文件。";
    return;
  }

  const supportedExtensions = [".docx", ".doc", ".rtf", ".txt", ".html", ".htm", ".odt"];
  const lowerName = file.name.toLowerCase();
  if (!supportedExtensions.some((extension) => lowerName.endsWith(extension))) {
    statusText.textContent = "当前支持 .docx、.doc、.rtf、.txt、.html、.odt。";
    return;
  }

  const data = new FormData();
  data.append("file", file);
  document.querySelectorAll(".format-options input, .format-options select").forEach((input) => {
    data.append(input.name, input.value);
  });
  submitButton.disabled = true;
  submitButton.textContent = "正在整理...";

  try {
    const response = await fetch("/api/format", {
      method: "POST",
      body: data,
    });
    const payload = await response.json();
    if (!response.ok) {
      throw new Error(payload.error || "处理失败");
    }

    downloadLink.href = payload.downloadUrl;
    downloadLink.download = payload.fileName;
    renderWarnings(payload.warnings || []);
    renderReport(payload.report || []);
    result.hidden = false;
    statusText.textContent = "";
  } catch (error) {
    statusText.textContent = error.message;
  } finally {
    submitButton.disabled = false;
    submitButton.textContent = "开始整理格式";
  }
});

function updateSelectedFile(file) {
  if (!file) {
    fileName.textContent = "选择或拖入 Word 文稿";
    fileMeta.textContent = "支持 .docx、.doc、.rtf、.txt、.html、.odt";
    dropzone.classList.remove("has-file");
    return;
  }

  fileName.textContent = file.name;
  fileMeta.textContent = `${formatFileSize(file.size)} · 处理后生成 .docx`;
  dropzone.classList.add("has-file");
}

function formatFileSize(bytes) {
  if (!Number.isFinite(bytes) || bytes <= 0) {
    return "文件已选择";
  }

  const units = ["B", "KB", "MB", "GB"];
  let size = bytes;
  let unitIndex = 0;
  while (size >= 1024 && unitIndex < units.length - 1) {
    size /= 1024;
    unitIndex += 1;
  }
  return `${size.toFixed(size >= 10 || unitIndex === 0 ? 0 : 1)} ${units[unitIndex]}`;
}

function renderWarnings(items) {
  warningList.innerHTML = "";
  if (!items.length) {
    warningList.innerHTML = '<p class="ok">未发现一级、二级标题编号跳号。</p>';
    return;
  }

  const fragment = document.createDocumentFragment();
  items.forEach((item) => {
    const row = document.createElement("div");
    row.className = "warning-row";
    row.innerHTML = `
      <strong>${escapeHtml(item.level)}</strong>
      <span>${escapeHtml(item.message)}</span>
    `;
    fragment.appendChild(row);
  });
  warningList.appendChild(fragment);
}

function renderReport(items) {
  reportList.innerHTML = "";
  if (!items.length) {
    reportList.innerHTML = '<p class="status">没有识别到可处理的正文段落。</p>';
    return;
  }

  const fragment = document.createDocumentFragment();
  items.slice(0, 40).forEach((item) => {
    const row = document.createElement("div");
    row.className = "report-row";
    row.innerHTML = `
      <span>第 ${item.index} 段</span>
      <strong>${escapeHtml(item.role)}</strong>
      <span>${escapeHtml(item.text)}</span>
    `;
    fragment.appendChild(row);
  });
  reportList.appendChild(fragment);
}

function escapeHtml(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

async function initializeOptions() {
  populateSizeSelects();
  const fonts = await loadSystemFonts();
  populateFontSelects(fonts);
}

async function loadSystemFonts() {
  try {
    const response = await fetch("/api/fonts");
    if (!response.ok) {
      throw new Error("无法读取字体");
    }
    const payload = await response.json();
    return Array.isArray(payload.fonts) && payload.fonts.length ? payload.fonts : fallbackFonts;
  } catch (_error) {
    return fallbackFonts;
  }
}

function populateFontSelects(fonts) {
  const uniqueFonts = Array.from(new Set([...fallbackFonts, ...fonts]))
    .filter(Boolean)
    .sort((a, b) => a.localeCompare(b, "zh-Hans-CN"));

  fontSelects.forEach((select) => {
    const defaultValue = select.dataset.default;
    select.innerHTML = "";
    if (!uniqueFonts.includes(defaultValue)) {
      uniqueFonts.unshift(defaultValue);
    }
    uniqueFonts.forEach((font) => {
      const option = document.createElement("option");
      option.value = font;
      option.textContent = font;
      select.appendChild(option);
    });
    select.value = defaultValue;
  });
}

function populateSizeSelects() {
  sizeSelects.forEach((select) => {
    const defaultValue = select.dataset.default;
    select.innerHTML = "";
    chineseFontSizes.forEach(([value, label]) => {
      const option = document.createElement("option");
      option.value = value;
      option.textContent = label;
      select.appendChild(option);
    });
    select.value = defaultValue;
  });
}
