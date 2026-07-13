const STORAGE_KEY = "excel-label-tool-v0.1";
const PAGE_SIZE = 20;

const state = {
  rows: [],
  currentPage: 0,
};

const els = {
  uploadPanel: document.getElementById("uploadPanel"),
  toolPanel: document.getElementById("toolPanel"),
  excelInput: document.getElementById("excelInput"),
  reuploadInput: document.getElementById("reuploadInput"),
  totalCount: document.getElementById("totalCount"),
  pageInfo: document.getElementById("pageInfo"),
  pageButtons: document.getElementById("pageButtons"),
  prevBtn: document.getElementById("prevBtn"),
  nextBtn: document.getElementById("nextBtn"),
  clearBtn: document.getElementById("clearBtn"),
  exportBtn: document.getElementById("exportBtn"),
  confirmNextBtn: document.getElementById("confirmNextBtn"),
  copyLlmBtn: document.getElementById("copyLlmBtn"),
  progressText: document.getElementById("progressText"),
  progressOk: document.getElementById("progressOk"),
  progressNeedsEdit: document.getElementById("progressNeedsEdit"),
  unconfirmedText: document.getElementById("unconfirmedText"),
  confirmedText: document.getElementById("confirmedText"),
  recordBody: document.getElementById("recordBody"),
  previewDialog: document.getElementById("previewDialog"),
  previewImage: document.getElementById("previewImage"),
  closePreviewBtn: document.getElementById("closePreviewBtn"),
  toast: document.getElementById("toast"),
};

function showToast(message) {
  els.toast.textContent = message;
  els.toast.classList.add("show");
  window.clearTimeout(showToast.timer);
  showToast.timer = window.setTimeout(() => {
    els.toast.classList.remove("show");
  }, 2200);
}

function normalizeHeader(value) {
  return String(value ?? "")
    .trim()
    .toLowerCase()
    .replace(/\s+/g, "");
}

function normalizeCell(value) {
  if (value === null || value === undefined) return "";
  return String(value).trim();
}

function isImageUrl(value) {
  const text = normalizeCell(value);
  return /^https?:\/\//i.test(text) || /^data:image\//i.test(text);
}

function findColumn(headers, candidates) {
  const normalizedCandidates = candidates.map(normalizeHeader);
  return headers.findIndex((header) => normalizedCandidates.includes(normalizeHeader(header)));
}

function parseRowsFromSheet(sheet) {
  const table = XLSX.utils.sheet_to_json(sheet, {
    header: 1,
    defval: "",
    raw: false,
  });
  const firstNonEmptyRowIndex = table.findIndex((row) => row.some((cell) => normalizeCell(cell)));
  if (firstNonEmptyRowIndex < 0) return [];

  const headers = table[firstNonEmptyRowIndex].map(normalizeCell);
  const idIndex = findColumn(headers, ["id", "编号", "序号"]);
  const originalIndex = findColumn(headers, ["原始", "原始文本", "original", "source", "输入"]);
  const llmIndex = findColumn(headers, ["llm", "识别文本", "模型结果", "识别结果", "结果"]);
  const resultIndex = findColumn(headers, ["修改结果", "人工结果", "final", "修正结果"]);
  const statusIndex = findColumn(headers, ["确认状态", "状态", "status"]);

  if (originalIndex < 0 || llmIndex < 0) {
    throw new Error("没有识别到“原始”和“llm”列，请检查表头。");
  }

  return table
    .slice(firstNonEmptyRowIndex + 1)
    .filter((row) => row.some((cell) => normalizeCell(cell)))
    .map((row, index) => {
      const llm = normalizeCell(row[llmIndex]);
      const existingResult = resultIndex >= 0 ? normalizeCell(row[resultIndex]) : "";
      const importedStatus = statusIndex >= 0 ? normalizeCell(row[statusIndex]) : "";
      return {
        id: normalizeCell(idIndex >= 0 ? row[idIndex] : "") || String(index + 1),
        original: normalizeCell(row[originalIndex]),
        llm,
        result: existingResult,
        status: parseStatus(importedStatus),
      };
    });
}

function parseStatus(value) {
  if (value.includes("无误") || value.toLowerCase() === "ok") return "ok";
  if (value.includes("修改") || value.toLowerCase() === "edit") return "edit";
  return "pending";
}

async function handleFile(file) {
  if (!file) return;
  if (!window.XLSX) {
    showToast("Excel 解析库未加载，请确认网络可访问后刷新页面。");
    return;
  }

  try {
    const buffer = await file.arrayBuffer();
    const workbook = XLSX.read(buffer, { type: "array" });
    const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
    const rows = parseRowsFromSheet(firstSheet);
    if (!rows.length) {
      showToast("没有读到有效数据。");
      return;
    }
    state.rows = rows;
    state.currentPage = 0;
    persist();
    render();
    showToast(`已导入 ${rows.length} 条数据`);
  } catch (error) {
    showToast(error.message || "Excel 读取失败。");
  } finally {
    els.excelInput.value = "";
    els.reuploadInput.value = "";
  }
}

function persist() {
  localStorage.setItem(
    STORAGE_KEY,
    JSON.stringify({
      rows: state.rows,
      currentPage: state.currentPage,
      savedAt: Date.now(),
    }),
  );
}

function restore() {
  try {
    const saved = JSON.parse(localStorage.getItem(STORAGE_KEY) || "{}");
    if (Array.isArray(saved.rows) && saved.rows.length) {
      state.rows = saved.rows;
      if (Number.isFinite(Number(saved.currentPage))) {
        state.currentPage = Number(saved.currentPage) || 0;
      } else {
        state.currentPage = Math.floor((Number(saved.currentIndex) || 0) / PAGE_SIZE);
      }
    }
  } catch {
    localStorage.removeItem(STORAGE_KEY);
  }
}

function statusLabel(row) {
  if (row.status === "ok") return "已确认无误";
  if (row.status === "edit") return "已确认需修改";
  return "未确认";
}

function statusClass(row) {
  if (row.status === "ok") return "ok";
  if (row.status === "edit") return "edit";
  return "";
}

function pageCount() {
  return Math.max(1, Math.ceil(state.rows.length / PAGE_SIZE));
}

function pageRange() {
  const start = state.currentPage * PAGE_SIZE;
  const end = Math.min(start + PAGE_SIZE, state.rows.length);
  return { start, end };
}

function visibleRows() {
  const { start, end } = pageRange();
  return state.rows.slice(start, end);
}

function pageButtonRange() {
  const totalPages = pageCount();
  const maxVisible = 20;
  if (totalPages <= maxVisible) {
    return { first: 0, last: totalPages - 1 };
  }

  const half = Math.floor(maxVisible / 2);
  let first = state.currentPage - half;
  let last = first + maxVisible - 1;

  if (first < 0) {
    first = 0;
    last = maxVisible - 1;
  }
  if (last >= totalPages) {
    last = totalPages - 1;
    first = totalPages - maxVisible;
  }

  return { first, last };
}

function renderContent(value) {
  const content = document.createElement("div");
  content.className = "content-box";
  if (isImageUrl(value)) {
    const image = document.createElement("img");
    image.src = value;
    image.alt = "图片内容";
    image.loading = "lazy";
    image.addEventListener("click", () => openPreview(value));
    image.addEventListener("error", () => {
      content.textContent = value;
      content.dataset.imageError = "true";
    });
    content.appendChild(image);
  } else {
    content.textContent = value || "-";
  }
  return content;
}

function render() {
  const hasRows = state.rows.length > 0;
  els.uploadPanel.classList.toggle("hidden", hasRows);
  els.toolPanel.classList.toggle("hidden", !hasRows);
  els.confirmNextBtn.disabled = !hasRows;

  if (!hasRows) {
    els.recordBody.innerHTML = "";
    els.pageButtons.innerHTML = "";
    updateProgress();
    return;
  }

  state.currentPage = Math.max(0, Math.min(state.currentPage, pageCount() - 1));
  const { start, end } = pageRange();

  els.totalCount.textContent = `总记录数：${state.rows.length}`;
  els.pageInfo.textContent = `第 ${state.currentPage + 1} / ${pageCount()} 页（${start + 1}-${end} / ${state.rows.length}）`;
  els.prevBtn.disabled = state.currentPage === 0;
  els.nextBtn.disabled = state.currentPage === pageCount() - 1;

  els.recordBody.innerHTML = "";
  visibleRows().forEach((row) => {
    els.recordBody.appendChild(renderRow(row));
  });
  renderPageButtons();
  updateProgress();
}

function renderPageButtons() {
  els.pageButtons.innerHTML = "";
  const totalPages = pageCount();
  if (!state.rows.length || totalPages <= 1) return;

  const { first, last } = pageButtonRange();
  for (let page = first; page <= last; page += 1) {
    const button = document.createElement("button");
    button.type = "button";
    button.className = `page-number ${page === state.currentPage ? "active" : ""}`;
    button.textContent = String(page + 1);
    button.disabled = page === state.currentPage;
    button.addEventListener("click", () => jumpToPage(page));
    els.pageButtons.appendChild(button);
  }
}

function renderRow(row) {
  const tr = document.createElement("tr");

  const idTd = document.createElement("td");
  idTd.className = "id-cell";
  const idNumber = document.createElement("div");
  idNumber.className = "id-number";
  idNumber.textContent = row.id;
  const status = document.createElement("div");
  status.className = `status-badge ${statusClass(row)}`;
  status.textContent = statusLabel(row);
  idTd.append(idNumber, status);

  const originalTd = document.createElement("td");
  originalTd.appendChild(renderContent(row.original));

  const llmTd = document.createElement("td");
  const llmTools = document.createElement("div");
  llmTools.className = "row-tools";
  const copyBtn = document.createElement("button");
  copyBtn.className = "mini-copy";
  copyBtn.type = "button";
  copyBtn.textContent = "复制";
  copyBtn.addEventListener("click", () => copyText(row.llm, "已复制本行 LLM 内容"));
  llmTools.appendChild(copyBtn);
  llmTd.append(llmTools, renderContent(row.llm));

  const resultTd = document.createElement("td");
  const textarea = document.createElement("textarea");
  textarea.className = "result-input";
  textarea.value = row.result || "";
  textarea.placeholder = "在这里填写或修改最终结果";
  textarea.addEventListener("input", () => {
    row.result = textarea.value;
    if (row.status !== "pending") {
      row.status = inferStatus(row);
    }
    persist();
    updateProgress();
    status.className = `status-badge ${statusClass(row)}`;
    status.textContent = statusLabel(row);
  });
  resultTd.appendChild(textarea);

  tr.append(idTd, originalTd, llmTd, resultTd);
  return tr;
}

function inferStatus(row) {
  return normalizeCell(row.result) === normalizeCell(row.llm) ? "ok" : "edit";
}

function updateProgress() {
  const total = state.rows.length;
  const okCount = state.rows.filter((row) => row.status === "ok").length;
  const editCount = state.rows.filter((row) => row.status === "edit").length;
  const confirmed = okCount + editCount;
  const unconfirmed = total - confirmed;
  const percent = total ? Math.round((confirmed / total) * 100) : 0;
  const okWidth = total ? (okCount / total) * 100 : 0;
  const editWidth = total ? (editCount / total) * 100 : 0;

  els.progressText.textContent = `${confirmed}/${total} (${percent}%)`;
  els.progressOk.style.width = `${okWidth}%`;
  els.progressNeedsEdit.style.left = `${okWidth}%`;
  els.progressNeedsEdit.style.width = `${editWidth}%`;
  els.unconfirmedText.textContent = `未确认：${unconfirmed}`;
  els.confirmedText.textContent = `已确认：${confirmed}`;
}

function confirmCurrentPageAndNext() {
  if (!state.rows.length) return;
  visibleRows().forEach((row) => {
    row.status = inferStatus(row);
  });
  if (state.currentPage < pageCount() - 1) {
    state.currentPage += 1;
    showToast("已确认本页，进入下一页");
  } else {
    showToast("已确认本页，已经是最后一页");
  }
  persist();
  render();
}

function go(delta) {
  if (!state.rows.length) return;
  state.currentPage = Math.max(0, Math.min(state.currentPage + delta, pageCount() - 1));
  persist();
  render();
}

function jumpToPage(page) {
  if (!state.rows.length) return;
  state.currentPage = Math.max(0, Math.min(page, pageCount() - 1));
  persist();
  render();
}

async function copyText(text, successMessage) {
  try {
    await navigator.clipboard.writeText(text || "");
    showToast(successMessage);
  } catch {
    showToast("复制失败，请手动选择内容复制。");
  }
}

function copyVisibleLlm() {
  if (!state.rows.length) return;
  const text = visibleRows()
    .map((row) => row.llm || "")
    .join("\n");
  copyText(text, "已复制本页 LLM 内容");
}

function clearData() {
  if (!state.rows.length) return;
  const ok = window.confirm("确定清除当前所有数据和标注进度吗？");
  if (!ok) return;
  state.rows = [];
  state.currentPage = 0;
  localStorage.removeItem(STORAGE_KEY);
  render();
  showToast("已清除数据");
}

function exportData() {
  if (!state.rows.length) {
    showToast("没有可导出的数据。");
    return;
  }
  if (!window.XLSX) {
    showToast("Excel 导出库未加载，请确认网络可访问后刷新页面。");
    return;
  }

  const data = state.rows.map((row) => ({
    id: row.id,
    原始: row.original,
    llm: row.llm,
    修改结果: row.result,
    确认状态: statusLabel(row),
  }));
  const sheet = XLSX.utils.json_to_sheet(data);
  const book = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(book, sheet, "标注结果");
  const timestamp = new Date().toISOString().slice(0, 19).replace(/[:T]/g, "-");
  XLSX.writeFile(book, `标注结果-${timestamp}.xlsx`);
}

function openPreview(src) {
  els.previewImage.src = src;
  if (typeof els.previewDialog.showModal === "function") {
    els.previewDialog.showModal();
  } else {
    window.open(src, "_blank", "noopener");
  }
}

function closePreview() {
  els.previewDialog.close();
  els.previewImage.removeAttribute("src");
}

els.excelInput.addEventListener("change", (event) => handleFile(event.target.files[0]));
els.reuploadInput.addEventListener("change", (event) => handleFile(event.target.files[0]));
els.prevBtn.addEventListener("click", () => go(-1));
els.nextBtn.addEventListener("click", () => go(1));
els.confirmNextBtn.addEventListener("click", confirmCurrentPageAndNext);
els.copyLlmBtn.addEventListener("click", copyVisibleLlm);
els.clearBtn.addEventListener("click", clearData);
els.exportBtn.addEventListener("click", exportData);
els.closePreviewBtn.addEventListener("click", closePreview);
els.previewDialog.addEventListener("click", (event) => {
  if (event.target === els.previewDialog) closePreview();
});

document.addEventListener("keydown", (event) => {
  if (event.altKey && event.key.toLowerCase() === "s") {
    event.preventDefault();
    confirmCurrentPageAndNext();
  }
});

restore();
render();
