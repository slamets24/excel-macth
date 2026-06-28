const els = {
  leftFile: document.getElementById("leftFile"),
  rightFile: document.getElementById("rightFile"),
  leftUploadText: document.getElementById("leftUploadText"),
  rightUploadText: document.getElementById("rightUploadText"),
  leftSummary: document.getElementById("leftSummary"),
  rightSummary: document.getElementById("rightSummary"),
  compareBtn: document.getElementById("compareBtn"),
  compareTopBtn: document.getElementById("compareTopBtn"),
  resetBtn: document.getElementById("resetBtn"),
  errorBox: document.getElementById("errorBox"),
  emptyState: document.getElementById("emptyState"),
  summaryGrid: document.getElementById("summaryGrid"),
  warningsBox: document.getElementById("warningsBox"),
  resultBox: document.getElementById("resultBox"),
  resultCount: document.getElementById("resultCount"),
  sharedCount: document.getElementById("sharedCount"),
  filters: document.getElementById("filters"),
  diffTable: document.getElementById("diffTable"),
  searchInput: document.getElementById("searchInput"),
  exportBtn: document.getElementById("exportBtn"),
};

const state = {
  left: null,
  right: null,
  rows: [],
  filteredRows: [],
  activeFilter: "all",
  search: "",
  summary: null,
};

const SIZE_ALIASES = new Map([
  ["XS", "XS"],
  ["S", "S"],
  ["M", "M"],
  ["L", "L"],
  ["XL", "XL"],
  ["XXL", "XXL"],
  ["2XL", "XXL"],
  ["XXXL", "XXXL"],
  ["3XL", "XXXL"],
  ["XXXXL", "XXXXL"],
  ["4XL", "XXXXL"],
]);

const SIZE_ORDER = ["XS", "S", "M", "L", "XL", "XXL", "XXXL", "XXXXL"];
const STATUS_META = {
  all: { label: "Semua", badge: "" },
  changed: { label: "Beda Qty", badge: "badge-changed" },
  added: { label: "Ada di File B", badge: "badge-added" },
  removed: { label: "Ada di File A", badge: "badge-removed" },
};

els.leftFile.addEventListener("change", () => handleFile("left", els.leftFile.files[0]));
els.rightFile.addEventListener("change", () => handleFile("right", els.rightFile.files[0]));
els.compareBtn.addEventListener("click", compareFiles);
els.compareTopBtn.addEventListener("click", compareFiles);
els.resetBtn.addEventListener("click", resetApp);
els.searchInput.addEventListener("input", () => {
  state.search = els.searchInput.value.trim().toLowerCase();
  applyFilters();
});
els.exportBtn.addEventListener("click", exportRows);

function showError(message) {
  els.errorBox.textContent = message;
  els.errorBox.style.display = message ? "block" : "none";
}

function normalizeText(value) {
  return String(value ?? "")
    .replace(/\s+/g, " ")
    .trim();
}

function normalizeKey(value) {
  return normalizeText(value).toUpperCase();
}

function numericValue(value) {
  if (typeof value === "number" && Number.isFinite(value)) return value;
  const cleaned = String(value ?? "")
    .replace(/\./g, "")
    .replace(",", ".")
    .replace(/[^\d.-]/g, "");
  if (!cleaned || cleaned === "-" || cleaned === ".") return 0;
  const parsed = Number(cleaned);
  return Number.isFinite(parsed) ? parsed : 0;
}

function formatNumber(value) {
  return new Intl.NumberFormat("id-ID", { maximumFractionDigits: 2 }).format(value || 0);
}

function formatDelta(value) {
  if (value > 0) return `+${formatNumber(value)}`;
  return formatNumber(value);
}

function isSize(value) {
  return SIZE_ALIASES.has(normalizeKey(value));
}

function normalizeSize(value) {
  return SIZE_ALIASES.get(normalizeKey(value)) || normalizeText(value);
}

async function handleFile(side, file) {
  showError("");
  if (!file) return;

  try {
    const buffer = await file.arrayBuffer();
    const workbook = XLSX.read(buffer, { type: "array", cellDates: true });
    const selectedSheet = workbook.SheetNames[0] || "";
    const parsed = parseWorkbook(workbook, selectedSheet, file.name);
    state[side] = { file, workbook, selectedSheet, parsed };
    updateUploadUI(side);
    updateCompareState();
  } catch (error) {
    state[side] = null;
    updateUploadUI(side);
    updateCompareState();
    showError(`Gagal membaca ${file.name}. Pastikan file Excel tidak rusak. Detail: ${error.message}`);
  }
}

function updateUploadUI(side) {
  const current = state[side];
  const summary = side === "left" ? els.leftSummary : els.rightSummary;
  const uploadText = side === "left" ? els.leftUploadText : els.rightUploadText;

  if (!current) {
    summary.style.display = "none";
    summary.innerHTML = "";
    uploadText.textContent = "Klik untuk upload Excel";
    return;
  }

  uploadText.textContent = current.file.name;
  const sheetOptions = current.workbook.SheetNames.map((name) => {
    const selected = name === current.selectedSheet ? " selected" : "";
    return `<option value="${escapeHtml(name)}"${selected}>${escapeHtml(name)}</option>`;
  }).join("");

  summary.style.display = "block";
  summary.innerHTML = `
    <p class="file-name" title="${escapeHtml(current.file.name)}">${escapeHtml(current.file.name)}</p>
    <div class="metrics">
      <div class="metric"><label>Mitra</label><strong>${formatNumber(current.parsed.partners.size)}</strong></div>
      <div class="metric"><label>Warna</label><strong>${formatNumber(current.parsed.colors.size)}</strong></div>
      <div class="metric"><label>Data</label><strong>${formatNumber(current.parsed.records.size)}</strong></div>
    </div>
    <label class="field-label" for="${side}Sheet">Sheet dipakai</label>
    <select id="${side}Sheet">${sheetOptions}</select>
  `;

  document.getElementById(`${side}Sheet`).addEventListener("change", (event) => {
    current.selectedSheet = event.target.value;
    try {
      current.parsed = parseWorkbook(current.workbook, current.selectedSheet, current.file.name);
      updateUploadUI(side);
      clearResults();
    } catch (error) {
      showError(`Gagal membaca sheet "${current.selectedSheet}" di ${current.file.name}. Detail: ${error.message}`);
    }
  });
}

function updateCompareState() {
  const ready = Boolean(state.left && state.right);
  els.compareBtn.disabled = !ready;
  els.compareTopBtn.disabled = !ready;
}

function parseWorkbook(workbook, sheetName, fileName) {
  const sheet = workbook.Sheets[sheetName];
  if (!sheet) throw new Error("Sheet tidak ditemukan.");

  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: "" });
  const layout = detectLayout(rows);
  const records = new Map();
  const partners = new Set();
  const colors = new Set();
  const warnings = [];

  if (!layout.columns.length) {
    throw new Error("Tidak menemukan kolom warna dan size. Pastikan header warna ada di atas size seperti XS, S, M, L.");
  }

  for (let rowIndex = layout.firstDataRow; rowIndex < rows.length; rowIndex += 1) {
    const row = rows[rowIndex] || [];
    const partner = normalizeText(row[layout.partnerCol]);
    if (!partner || isIgnoredPartner(partner)) continue;

    let rowTotal = 0;
    for (const col of layout.columns) {
      rowTotal += Math.abs(numericValue(row[col.index]));
    }
    if (rowTotal === 0 && !hasAnyTextMetric(row, layout.columns)) continue;

    partners.add(normalizeKey(partner));

    for (const col of layout.columns) {
      const qty = numericValue(row[col.index]);
      if (qty === 0) continue;
      const color = normalizeText(col.color);
      const size = normalizeSize(col.size);
      const key = makeRecordKey(partner, color, size);
      colors.add(normalizeKey(color));

      if (!records.has(key)) {
        records.set(key, { partner, color, size, qty: 0 });
      }
      records.get(key).qty += qty;
    }
  }

  if (!records.size) {
    warnings.push(`${fileName}: sheet "${sheetName}" terbaca, tapi tidak ada qty mitra/warna/size yang bernilai.`);
  }

  return { fileName, sheetName, rows, layout, records, partners, colors, warnings };
}

function detectLayout(rows) {
  const scanLimit = Math.min(rows.length, 30);
  let partnerRow = -1;
  let partnerCol = -1;

  for (let r = 0; r < scanLimit; r += 1) {
    const row = rows[r] || [];
    for (let c = 0; c < Math.min(row.length, 20); c += 1) {
      const value = normalizeKey(row[c]);
      if (["NAMA TOKO", "NAMA MITRA", "MITRA", "TOKO", "CUSTOMER"].includes(value)) {
        partnerRow = r;
        partnerCol = c;
        break;
      }
    }
    if (partnerCol >= 0) break;
  }

  if (partnerCol < 0) {
    throw new Error("Kolom mitra tidak ditemukan. Header yang dicari: NAMA TOKO atau NAMA MITRA.");
  }

  let sizeRow = -1;
  let bestScore = 0;
  for (let r = Math.max(0, partnerRow - 2); r < scanLimit; r += 1) {
    const row = rows[r] || [];
    const score = row.reduce((count, value, index) => {
      if (index <= partnerCol) return count;
      return count + (isSize(value) ? 1 : 0);
    }, 0);
    if (score > bestScore) {
      bestScore = score;
      sizeRow = r;
    }
  }

  if (sizeRow < 0 || bestScore < 2) {
    throw new Error("Baris size tidak ditemukan. Minimal perlu dua header size seperti XS, S, M, L.");
  }

  const colorRow = Math.max(0, sizeRow - 1);
  const maxCols = Math.max(...rows.slice(0, scanLimit).map((row) => row.length), 0);
  const columns = [];
  let currentColor = "";

  for (let c = partnerCol + 1; c < maxCols; c += 1) {
    const colorCandidate = normalizeText(rows[colorRow]?.[c]);
    const sizeCandidate = normalizeText(rows[sizeRow]?.[c]);
    if (normalizeKey(colorCandidate) === "TOTAL" || normalizeKey(sizeCandidate) === "TOTAL") break;
    if (colorCandidate && !isSize(colorCandidate)) currentColor = colorCandidate;
    if (currentColor && isSize(sizeCandidate)) {
      columns.push({ index: c, color: currentColor, size: normalizeSize(sizeCandidate) });
    }
  }

  let firstDataRow = sizeRow + 1;
  while (firstDataRow < rows.length) {
    const partner = normalizeText(rows[firstDataRow]?.[partnerCol]);
    const hasMetric = columns.some((col) => numericValue(rows[firstDataRow]?.[col.index]) !== 0);
    if (partner && !isIgnoredPartner(partner) && hasMetric) break;
    firstDataRow += 1;
  }

  return { partnerRow, partnerCol, colorRow, sizeRow, firstDataRow, columns };
}

function isIgnoredPartner(value) {
  const key = normalizeKey(value);
  return ["NAMA TOKO", "NAMA MITRA", "MITRA", "TOKO", "TOTAL", "GRAND TOTAL", "JUMLAH"].includes(key);
}

function hasAnyTextMetric(row, columns) {
  return columns.some((col) => normalizeText(row[col.index]));
}

function makeRecordKey(partner, color, size) {
  return `${normalizeKey(partner)}||${normalizeKey(color)}||${normalizeKey(size)}`;
}

function compareFiles() {
  showError("");
  if (!state.left || !state.right) {
    showError("Upload File A dan File B dulu.");
    return;
  }

  const left = state.left.parsed;
  const right = state.right.parsed;
  const keys = new Set([...left.records.keys(), ...right.records.keys()]);
  const rows = [];

  for (const key of keys) {
    const leftRecord = left.records.get(key);
    const rightRecord = right.records.get(key);
    const base = rightRecord || leftRecord;
    const leftQty = leftRecord?.qty || 0;
    const rightQty = rightRecord?.qty || 0;
    if (leftQty === rightQty) continue;

    const delta = rightQty - leftQty;
    let status = "changed";
    if (leftQty === 0 && rightQty !== 0) status = "added";
    if (leftQty !== 0 && rightQty === 0) status = "removed";

    rows.push({
      partner: base.partner,
      color: base.color,
      size: base.size,
      leftQty,
      rightQty,
      delta,
      status,
    });
  }

  rows.sort(sortDiffRows);
  state.rows = rows;
  state.summary = buildSummary(left, right, rows);
  state.activeFilter = "all";
  state.search = "";
  els.searchInput.value = "";
  renderResults();
}

function sortDiffRows(a, b) {
  const byPartner = a.partner.localeCompare(b.partner, "id", { sensitivity: "base" });
  if (byPartner) return byPartner;
  const byColor = a.color.localeCompare(b.color, "id", { sensitivity: "base" });
  if (byColor) return byColor;
  return SIZE_ORDER.indexOf(a.size) - SIZE_ORDER.indexOf(b.size);
}

function buildSummary(left, right, rows) {
  const sharedPartners = [...left.partners].filter((partner) => right.partners.has(partner)).length;
  const totalDelta = rows.reduce((sum, row) => sum + Math.abs(row.delta), 0);
  const changedPartners = new Set(rows.map((row) => normalizeKey(row.partner))).size;

  return {
    leftPartners: left.partners.size,
    rightPartners: right.partners.size,
    sharedPartners,
    changedPartners,
    diffRows: rows.length,
    totalDelta,
    warnings: [...left.warnings, ...right.warnings],
  };
}

function renderResults() {
  const summary = state.summary;
  els.emptyState.style.display = "none";
  els.resultBox.style.display = "block";
  els.summaryGrid.style.display = "grid";
  els.sharedCount.textContent = formatNumber(summary.sharedPartners);
  els.summaryGrid.innerHTML = `
    ${summaryCard("Mitra File A", summary.leftPartners)}
    ${summaryCard("Mitra File B", summary.rightPartners)}
    ${summaryCard("Mitra berbeda", summary.changedPartners, "summary-amber")}
    ${summaryCard("Warna/Size beda", summary.diffRows, "summary-red")}
    ${summaryCard("Total selisih qty", summary.totalDelta, "summary-green")}
  `;

  els.warningsBox.style.display = summary.warnings.length ? "block" : "none";
  els.warningsBox.innerHTML = summary.warnings.map(escapeHtml).join("<br>");
  renderFilters();
  applyFilters();
}

function summaryCard(label, value, extraClass = "") {
  return `<div class="summary-card ${extraClass}"><label>${escapeHtml(label)}</label><strong>${formatNumber(value)}</strong></div>`;
}

function renderFilters() {
  const counts = state.rows.reduce((acc, row) => {
    acc[row.status] = (acc[row.status] || 0) + 1;
    return acc;
  }, { all: state.rows.length });

  els.filters.innerHTML = ["all", "changed", "added", "removed"].map((status) => {
    const active = state.activeFilter === status ? " active" : "";
    return `<button class="filter-btn${active}" type="button" data-filter="${status}">${STATUS_META[status].label} (${formatNumber(counts[status] || 0)})</button>`;
  }).join("");

  els.filters.querySelectorAll("button").forEach((button) => {
    button.addEventListener("click", () => {
      state.activeFilter = button.dataset.filter;
      renderFilters();
      applyFilters();
    });
  });
}

function applyFilters() {
  const query = state.search;
  state.filteredRows = state.rows.filter((row) => {
    if (state.activeFilter !== "all" && row.status !== state.activeFilter) return false;
    if (!query) return true;
    return [row.partner, row.color, row.size].some((value) => normalizeText(value).toLowerCase().includes(query));
  });
  renderTable();
}

function renderTable() {
  els.resultCount.textContent = `Menampilkan ${formatNumber(state.filteredRows.length)} dari ${formatNumber(state.rows.length)} baris berbeda`;

  if (!state.filteredRows.length) {
    els.diffTable.innerHTML = `
      <thead><tr><th>Mitra</th><th>Status</th><th>Warna</th><th>Size</th><th>File A</th><th>File B</th><th>Selisih</th></tr></thead>
      <tbody><tr><td colspan="7">Tidak ada hasil untuk filter ini.</td></tr></tbody>
    `;
    return;
  }

  const body = state.filteredRows.map((row) => {
    const deltaClass = row.delta > 0 ? "num-pos" : row.delta < 0 ? "num-neg" : "";
    return `
      <tr class="row-${row.status}">
        <td class="partner">${escapeHtml(row.partner)}</td>
        <td><span class="badge ${STATUS_META[row.status].badge}">${STATUS_META[row.status].label}</span></td>
        <td>${escapeHtml(row.color)}</td>
        <td>${escapeHtml(row.size)}</td>
        <td>${formatNumber(row.leftQty)}</td>
        <td>${formatNumber(row.rightQty)}</td>
        <td class="${deltaClass}">${formatDelta(row.delta)}</td>
      </tr>
    `;
  }).join("");

  els.diffTable.innerHTML = `
    <thead>
      <tr>
        <th>Mitra</th>
        <th>Status</th>
        <th>Warna</th>
        <th>Size</th>
        <th>File A</th>
        <th>File B</th>
        <th>Selisih</th>
      </tr>
    </thead>
    <tbody>${body}</tbody>
  `;
}

function exportRows() {
  if (!state.rows.length) {
    showError("Belum ada hasil berbeda untuk diexport.");
    return;
  }

  const data = state.filteredRows.map((row) => ({
    Mitra: row.partner,
    Status: STATUS_META[row.status].label,
    Warna: row.color,
    Size: row.size,
    "File A": row.leftQty,
    "File B": row.rightQty,
    Selisih: row.delta,
  }));
  const worksheet = XLSX.utils.json_to_sheet(data);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "Hasil Compare");
  XLSX.writeFile(workbook, "excel-match-mitra-warna-size.xlsx");
}

function clearResults() {
  state.rows = [];
  state.filteredRows = [];
  state.summary = null;
  els.emptyState.style.display = "block";
  els.resultBox.style.display = "none";
  els.summaryGrid.style.display = "none";
  els.warningsBox.style.display = "none";
  els.sharedCount.textContent = "0";
}

function resetApp() {
  state.left = null;
  state.right = null;
  state.activeFilter = "all";
  state.search = "";
  els.leftFile.value = "";
  els.rightFile.value = "";
  els.searchInput.value = "";
  updateUploadUI("left");
  updateUploadUI("right");
  updateCompareState();
  clearResults();
  showError("");
}

function escapeHtml(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}
