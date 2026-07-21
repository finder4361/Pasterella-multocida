/* Pasteurella Taiwan Dashboard
 * Static browser app: Excel/CSV parsing, Google Sheets CSV loading, deduplication,
 * filtering, Plotly charts, and privacy-aware table output.
 */

const state = {
  workbook: null,
  rawRows: [],
  normalizedRows: [],
  filteredRows: [],
  activeSheetName: null,
};

const FIELD_ALIASES = {
  include: ["納入年度統計", "納入統計", "有資料"],
  yearAD: ["西元年", "年度", "year", "Year"],
  yearROC: ["民國年", "ROC年"],
  region: ["病例區域", "區域", "主要病例區域"],
  county: ["原始縣市", "縣市", "場址縣市"],
  township: ["原始鄉鎮", "鄉鎮", "場址鄉鎮"],
  classification: ["區域屬性", "分類性質"],
  isolateSite: ["分離部位", "採樣部位"],
  immunologyId: ["免疫室編號", "檢驗編號", "病例編號"],
  owner: ["畜主", "畜主姓名", "豬場", "場戶"],
  dedupKey: ["去重鍵", "唯一鍵", "UniqueID", "ID"],
  duplicateTube: ["是否重複保管管", "重複保管管"],
  quality: ["資料品質註記", "品質註記", "備註"],
  positiveCount: ["陽性數", "分離陽性數", "陽性病例數", "positive_count"],
  denominator: ["總檢體數", "總送檢數", "病例總數", "受檢數", "denominator", "total_count"],
};

const els = {
  fileInput: document.getElementById("fileInput"),
  sheetUrlInput: document.getElementById("sheetUrlInput"),
  loadSheetButton: document.getElementById("loadSheetButton"),
  loadDemoButton: document.getElementById("loadDemoButton"),
  sheetSelect: document.getElementById("sheetSelect"),
  metricSelect: document.getElementById("metricSelect"),
  yearSelect: document.getElementById("yearSelect"),
  regionSelect: document.getElementById("regionSelect"),
  countySelect: document.getElementById("countySelect"),
  includePendingCheckbox: document.getElementById("includePendingCheckbox"),
  showOwnerCheckbox: document.getElementById("showOwnerCheckbox"),
  statusMessage: document.getElementById("statusMessage"),
  downloadCsvButton: document.getElementById("downloadCsvButton"),
  tableBody: document.querySelector("#dataTable tbody"),
  ownerHeader: document.querySelector("th.owner-column"),
};

function valueFromAliases(row, aliases) {
  for (const key of aliases) {
    if (Object.prototype.hasOwnProperty.call(row, key)) {
      const value = row[key];
      if (value !== null && value !== undefined && String(value).trim() !== "") return value;
    }
  }
  return "";
}

function truthyChinese(value) {
  const normalized = String(value ?? "").trim().toLowerCase();
  if (normalized === "") return true;
  return !["否", "false", "0", "不納入", "no"].includes(normalized);
}

function numberOrNull(value) {
  if (value === null || value === undefined || value === "") return null;
  const cleaned = String(value).replace(/,/g, "").replace(/%/g, "").trim();
  const parsed = Number(cleaned);
  return Number.isFinite(parsed) ? parsed : null;
}

function normalizeYear(row) {
  const ad = numberOrNull(valueFromAliases(row, FIELD_ALIASES.yearAD));
  if (ad && ad >= 1900 && ad <= 2200) return Math.round(ad);
  const roc = numberOrNull(valueFromAliases(row, FIELD_ALIASES.yearROC));
  if (roc && roc > 0 && roc < 300) return Math.round(roc + 1911);
  return null;
}

function normalizeRow(row, index) {
  const year = normalizeYear(row);
  const immunologyId = String(valueFromAliases(row, FIELD_ALIASES.immunologyId) || "").trim();
  const isolateSite = String(valueFromAliases(row, FIELD_ALIASES.isolateSite) || "").trim();
  const explicitKey = String(valueFromAliases(row, FIELD_ALIASES.dedupKey) || "").trim();
  const fallbackKey = [year ?? "NA", immunologyId || "NA", isolateSite || "NA"].join("|");

  return {
    _sourceIndex: index + 2,
    _raw: row,
    include: truthyChinese(valueFromAliases(row, FIELD_ALIASES.include)),
    year,
    region: String(valueFromAliases(row, FIELD_ALIASES.region) || "待確認").trim() || "待確認",
    county: String(valueFromAliases(row, FIELD_ALIASES.county) || "待確認").trim() || "待確認",
    township: String(valueFromAliases(row, FIELD_ALIASES.township) || "").trim(),
    classification: String(valueFromAliases(row, FIELD_ALIASES.classification) || "待確認").trim() || "待確認",
    isolateSite,
    immunologyId,
    owner: String(valueFromAliases(row, FIELD_ALIASES.owner) || "").trim(),
    dedupKey: explicitKey || fallbackKey,
    duplicateTube: ["是", "true", "1", "yes"].includes(
      String(valueFromAliases(row, FIELD_ALIASES.duplicateTube) || "").trim().toLowerCase()
    ),
    quality: String(valueFromAliases(row, FIELD_ALIASES.quality) || "").trim(),
    positiveCount: numberOrNull(valueFromAliases(row, FIELD_ALIASES.positiveCount)),
    denominator: numberOrNull(valueFromAliases(row, FIELD_ALIASES.denominator)),
  };
}

function deduplicateRows(rows) {
  const seen = new Set();
  return rows.filter((row) => {
    if (!row.include || row.duplicateTube) return false;
    if (seen.has(row.dedupKey)) return false;
    seen.add(row.dedupKey);
    return true;
  });
}

function setStatus(message, type = "info") {
  els.statusMessage.className = `status ${type}`;
  els.statusMessage.textContent = message;
}

function convertGoogleSheetUrl(input) {
  const value = input.trim();
  if (!value) throw new Error("請貼上 Google 試算表網址。");
  if (/output=csv|tqx=out:csv|format=csv/i.test(value)) return value;

  const idMatch = value.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
  if (!idMatch) return value;

  const gidMatch = value.match(/[?&#]gid=(\d+)/);
  const gid = gidMatch ? gidMatch[1] : "0";
  return `https://docs.google.com/spreadsheets/d/${idMatch[1]}/gviz/tq?tqx=out:csv&gid=${gid}`;
}

function rowsFromWorksheet(ws) {
  return XLSX.utils.sheet_to_json(ws, { defval: "", raw: false });
}

function chooseDefaultSheet(sheetNames) {
  const preferences = ["區域分類明細", "年度地區分離率", "整併資料", "分析摘要"];
  return preferences.find((name) => sheetNames.includes(name)) || sheetNames[0];
}

function populateSheetSelect(sheetNames, selected) {
  els.sheetSelect.innerHTML = "";
  for (const name of sheetNames) {
    const option = new Option(name, name, name === selected, name === selected);
    els.sheetSelect.add(option);
  }
  els.sheetSelect.disabled = false;
}

function loadWorkbookSheet(sheetName) {
  if (!state.workbook || !state.workbook.Sheets[sheetName]) return;
  state.activeSheetName = sheetName;
  state.rawRows = rowsFromWorksheet(state.workbook.Sheets[sheetName]);
  processRows(state.rawRows, `已載入工作表「${sheetName}」`);
}

function processRows(rawRows, label) {
  state.rawRows = rawRows;
  state.normalizedRows = deduplicateRows(rawRows.map(normalizeRow));
  populateFilters();
  applyFiltersAndRender();

  const hasRate = state.normalizedRows.some(
    (r) => r.positiveCount !== null && r.denominator !== null && r.denominator > 0
  );
  setStatus(
    `${label}：原始 ${rawRows.length.toLocaleString()} 列；納入並去重後 ${state.normalizedRows.length.toLocaleString()} 筆。` +
      (hasRate ? " 已偵測到陽性數與分母，可顯示真正分離率。" : " 未偵測到分母，目前「分離率」模式會提示缺少資料。"),
    hasRate ? "success" : "warning"
  );
}

function populateFilters() {
  const years = [...new Set(state.normalizedRows.map((r) => r.year).filter(Boolean))].sort((a, b) => a - b);
  const regions = [...new Set(state.normalizedRows.map((r) => r.region).filter(Boolean))].sort();
  const counties = [...new Set(state.normalizedRows.map((r) => r.county).filter(Boolean))].sort();

  fillSelect(els.yearSelect, years, "ALL", "全部年度");
  fillSelect(els.regionSelect, regions, "ALL", "全部區域");
  fillSelect(els.countySelect, counties, "ALL", "全部縣市");
}

function fillSelect(select, values, defaultValue, defaultLabel) {
  const current = select.value;
  select.innerHTML = "";
  select.add(new Option(defaultLabel, defaultValue));
  values.forEach((value) => select.add(new Option(String(value), String(value))));
  select.value = values.map(String).includes(current) ? current : defaultValue;
}

function applyFilters() {
  const year = els.yearSelect.value;
  const region = els.regionSelect.value;
  const county = els.countySelect.value;
  const includePending = els.includePendingCheckbox.checked;

  state.filteredRows = state.normalizedRows.filter((row) => {
    if (year !== "ALL" && String(row.year) !== year) return false;
    if (region !== "ALL" && row.region !== region) return false;
    if (county !== "ALL" && row.county !== county) return false;
    if (!includePending && (row.region === "待確認" || row.county === "待確認")) return false;
    return true;
  });
}

function applyFiltersAndRender() {
  applyFilters();
  renderKpis();
  renderCharts();
  renderTable();
}

function renderKpis() {
  const rows = state.filteredRows;
  const years = rows.map((r) => r.year).filter(Boolean);
  const pending = rows.filter((r) => r.region === "待確認" || r.county === "待確認").length;
  const resolvedPct = rows.length ? ((rows.length - pending) / rows.length) * 100 : 0;

  document.getElementById("kpiRecords").textContent = rows.length.toLocaleString();
  document.getElementById("kpiYears").textContent = years.length
    ? `${Math.min(...years)}–${Math.max(...years)}`
    : "—";
  document.getElementById("kpiResolved").textContent = rows.length ? `${resolvedPct.toFixed(1)}%` : "—";
  document.getElementById("kpiPending").textContent = pending.toLocaleString();
}

function groupBy(rows, keyFn) {
  const map = new Map();
  rows.forEach((row) => {
    const key = keyFn(row);
    if (!map.has(key)) map.set(key, []);
    map.get(key).push(row);
  });
  return map;
}

function metricValue(rows, metric) {
  if (metric === "rate") {
    const valid = rows.filter((r) => r.positiveCount !== null && r.denominator !== null && r.denominator > 0);
    const positive = valid.reduce((sum, r) => sum + r.positiveCount, 0);
    const denominator = valid.reduce((sum, r) => sum + r.denominator, 0);
    return denominator ? (positive / denominator) * 100 : null;
  }
  return rows.length;
}

function renderCharts() {
  const rows = state.filteredRows;
  const metric = els.metricSelect.value;
  renderAnnualChart(rows, metric);
  renderCountyChart(rows, metric);
  renderClassificationChart(rows);
  renderQualityChart(rows);
}

function baseLayout() {
  return {
    font: { family: 'Inter, "Noto Sans TC", "Microsoft JhengHei", sans-serif' },
    margin: { l: 60, r: 20, t: 28, b: 55 },
    paper_bgcolor: "rgba(0,0,0,0)",
    plot_bgcolor: "rgba(0,0,0,0)",
    hovermode: "closest",
    legend: { orientation: "h", y: 1.12 },
  };
}

function renderEmptyChart(id, message) {
  Plotly.react(
    id,
    [],
    {
      ...baseLayout(),
      annotations: [{ text: message, showarrow: false, x: 0.5, y: 0.5, xref: "paper", yref: "paper" }],
      xaxis: { visible: false },
      yaxis: { visible: false },
    },
    { responsive: true, displaylogo: false }
  );
}

function renderAnnualChart(rows, metric) {
  if (!rows.length) return renderEmptyChart("annualChart", "目前篩選條件沒有資料");

  const years = [...new Set(rows.map((r) => r.year).filter(Boolean))].sort((a, b) => a - b);
  const regions = [...new Set(rows.map((r) => r.region))].sort();
  const metricName = metric === "rate" ? "分離率 (%)" : metric === "share" ? "年度內菌株比例 (%)" : "去重分離紀錄數";

  if (metric === "rate" && !rows.some((r) => r.denominator && r.positiveCount !== null)) {
    return renderEmptyChart("annualChart", "缺少「陽性數」與「總檢體數／總送檢數」，無法計算分離率");
  }

  const traces = regions.map((region) => {
    const values = years.map((year) => {
      const subset = rows.filter((r) => r.year === year && r.region === region);
      if (metric === "share") {
        const total = rows.filter((r) => r.year === year).length;
        return total ? (subset.length / total) * 100 : 0;
      }
      return metricValue(subset, metric);
    });
    return { type: "bar", name: region, x: years, y: values, hovertemplate: `%{x}<br>${region}: %{y:.2f}<extra></extra>` };
  });

  document.getElementById("annualChartNote").textContent =
    metric === "rate"
      ? "分離率＝陽性數 ÷ 總檢體數；圖中以百分比呈現。"
      : metric === "share"
      ? "每一年度各區域菌株數除以該年度全部菌株數。"
      : "去除重複保管管後的分離紀錄數。";

  Plotly.react(
    "annualChart",
    traces,
    {
      ...baseLayout(),
      barmode: metric === "share" ? "stack" : "group",
      xaxis: { title: "西元年", type: "category" },
      yaxis: { title: metricName, rangemode: "tozero", ticksuffix: metric === "count" ? "" : "%" },
    },
    { responsive: true, displaylogo: false }
  );
}

function renderCountyChart(rows, metric) {
  if (!rows.length) return renderEmptyChart("countyChart", "目前篩選條件沒有資料");
  if (metric === "rate" && !rows.some((r) => r.denominator && r.positiveCount !== null)) {
    return renderEmptyChart("countyChart", "缺少分母，無法計算縣市分離率");
  }

  const grouped = groupBy(rows, (r) => r.county);
  let data = [...grouped.entries()].map(([county, subset]) => {
    let value = metricValue(subset, metric);
    if (metric === "share") value = rows.length ? (subset.length / rows.length) * 100 : 0;
    return { county, value };
  }).filter((d) => d.value !== null);

  data.sort((a, b) => b.value - a.value);
  data = data.slice(0, 15).reverse();

  Plotly.react(
    "countyChart",
    [{
      type: "bar",
      orientation: "h",
      x: data.map((d) => d.value),
      y: data.map((d) => d.county),
      hovertemplate: "%{y}: %{x:.2f}<extra></extra>",
    }],
    {
      ...baseLayout(),
      margin: { l: 95, r: 20, t: 25, b: 55 },
      xaxis: { title: metric === "count" ? "去重分離紀錄數" : "百分比 (%)", rangemode: "tozero" },
      yaxis: { automargin: true },
      showlegend: false,
    },
    { responsive: true, displaylogo: false }
  );
}

function renderClassificationChart(rows) {
  if (!rows.length) return renderEmptyChart("classificationChart", "目前篩選條件沒有資料");
  const grouped = groupBy(rows, (r) => r.classification || "待確認");
  const labels = [...grouped.keys()];
  const values = labels.map((label) => grouped.get(label).length);

  Plotly.react(
    "classificationChart",
    [{
      type: "pie",
      labels,
      values,
      hole: 0.58,
      textinfo: "label+percent",
      hovertemplate: "%{label}<br>%{value} 筆<br>%{percent}<extra></extra>",
    }],
    { ...baseLayout(), margin: { l: 15, r: 15, t: 25, b: 25 }, showlegend: false },
    { responsive: true, displaylogo: false }
  );
}

function renderQualityChart(rows) {
  if (!rows.length) return renderEmptyChart("qualityChart", "目前篩選條件沒有資料");
  const years = [...new Set(rows.map((r) => r.year).filter(Boolean))].sort((a, b) => a - b);
  const resolved = years.map((year) => rows.filter((r) => r.year === year && r.region !== "待確認" && r.county !== "待確認").length);
  const pending = years.map((year) => rows.filter((r) => r.year === year && (r.region === "待確認" || r.county === "待確認")).length);

  Plotly.react(
    "qualityChart",
    [
      { type: "scatter", mode: "lines+markers", name: "地區已判定", x: years, y: resolved },
      { type: "scatter", mode: "lines+markers", name: "待確認", x: years, y: pending },
    ],
    {
      ...baseLayout(),
      xaxis: { title: "西元年", type: "category" },
      yaxis: { title: "紀錄數", rangemode: "tozero" },
    },
    { responsive: true, displaylogo: false }
  );
}

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function renderTable() {
  const showOwner = els.showOwnerCheckbox.checked;
  els.ownerHeader.classList.toggle("hidden", !showOwner);
  document.querySelectorAll("td.owner-column").forEach((td) => td.classList.toggle("hidden", !showOwner));

  const rows = state.filteredRows.slice(0, 500);
  if (!rows.length) {
    els.tableBody.innerHTML = `<tr><td colspan="${showOwner ? 8 : 7}" class="empty-cell">目前篩選條件沒有資料</td></tr>`;
    return;
  }

  els.tableBody.innerHTML = rows.map((row) => `
    <tr>
      <td>${escapeHtml(row.year ?? "")}</td>
      <td>${escapeHtml(row.region)}</td>
      <td>${escapeHtml(row.county)}</td>
      <td>${escapeHtml(row.classification)}</td>
      <td>${escapeHtml(row.isolateSite)}</td>
      <td>${escapeHtml(row.immunologyId)}</td>
      <td class="owner-column ${showOwner ? "" : "hidden"}">${escapeHtml(row.owner)}</td>
      <td>${escapeHtml(row.quality)}</td>
    </tr>
  `).join("");
}

function downloadFilteredCsv() {
  if (!state.filteredRows.length) {
    setStatus("目前沒有可下載的篩選結果。", "error");
    return;
  }
  const showOwner = els.showOwnerCheckbox.checked;
  const data = state.filteredRows.map((row) => ({
    西元年: row.year,
    病例區域: row.region,
    原始縣市: row.county,
    原始鄉鎮: row.township,
    區域屬性: row.classification,
    分離部位: row.isolateSite,
    免疫室編號: row.immunologyId,
    ...(showOwner ? { 畜主: row.owner } : {}),
    資料品質註記: row.quality,
  }));
  const csv = Papa.unparse(data);
  const blob = new Blob(["\ufeff", csv], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = `pasteurella_filtered_${new Date().toISOString().slice(0, 10)}.csv`;
  a.click();
  URL.revokeObjectURL(url);
}

async function loadLocalFile(file) {
  if (!file) return;
  try {
    const buffer = await file.arrayBuffer();
    if (file.name.toLowerCase().endsWith(".csv")) {
      const text = new TextDecoder("utf-8").decode(buffer);
      const parsed = Papa.parse(text, { header: true, skipEmptyLines: true });
      if (parsed.errors.length) console.warn(parsed.errors);
      state.workbook = null;
      els.sheetSelect.innerHTML = `<option>CSV</option>`;
      els.sheetSelect.disabled = true;
      processRows(parsed.data, `已載入 ${file.name}`);
      return;
    }

    state.workbook = XLSX.read(buffer, { type: "array" });
    const defaultSheet = chooseDefaultSheet(state.workbook.SheetNames);
    populateSheetSelect(state.workbook.SheetNames, defaultSheet);
    loadWorkbookSheet(defaultSheet);
  } catch (error) {
    console.error(error);
    setStatus(`讀取檔案失敗：${error.message}`, "error");
  }
}

async function loadGoogleSheet() {
  try {
    const url = convertGoogleSheetUrl(els.sheetUrlInput.value);
    setStatus("正在讀取公開試算表資料…", "info");
    const response = await fetch(url);
    if (!response.ok) throw new Error(`HTTP ${response.status}`);
    const text = await response.text();
    const parsed = Papa.parse(text, { header: true, skipEmptyLines: true });
    if (!parsed.data.length) throw new Error("CSV 沒有可讀取的資料列");
    state.workbook = null;
    els.sheetSelect.innerHTML = `<option>Google Sheets</option>`;
    els.sheetSelect.disabled = true;
    processRows(parsed.data, "已載入 Google 試算表");
  } catch (error) {
    console.error(error);
    setStatus(
      `Google 試算表讀取失敗：${error.message}。請確認檔案已設為公開或「發布到網路」，並使用正確 gid。`,
      "error"
    );
  }
}

function loadDemo() {
  const demoRows = [
    { 西元年: 2021, 免疫室編號: "DEMO-001", 分離部位: "LC", 病例區域: "雲林區", 原始縣市: "雲林縣", 區域屬性: "核心縣市", 去重鍵: "2021-001-LC" },
    { 西元年: 2021, 免疫室編號: "DEMO-002", 分離部位: "RLC", 病例區域: "高雄屏東區", 原始縣市: "屏東縣", 區域屬性: "核心縣市", 去重鍵: "2021-002-RLC" },
    { 西元年: 2022, 免疫室編號: "DEMO-003", 分離部位: "LD", 病例區域: "雲林區", 原始縣市: "彰化縣", 區域屬性: "鄰近歸區", 去重鍵: "2022-003-LD" },
    { 西元年: 2022, 免疫室編號: "DEMO-004", 分離部位: "LLD", 病例區域: "高雄屏東區", 原始縣市: "高雄市", 區域屬性: "核心縣市", 去重鍵: "2022-004-LLD" },
    { 西元年: 2023, 免疫室編號: "DEMO-005", 分離部位: "LU", 病例區域: "待確認", 原始縣市: "待確認", 區域屬性: "待確認", 去重鍵: "2023-005-LU", 資料品質註記: "格式示範資料，非研究結果" },
  ];
  state.workbook = null;
  els.sheetSelect.innerHTML = `<option>格式示範</option>`;
  els.sheetSelect.disabled = true;
  processRows(demoRows, "已載入格式示範（非真實研究數據）");
}

els.fileInput.addEventListener("change", (event) => loadLocalFile(event.target.files[0]));
els.loadSheetButton.addEventListener("click", loadGoogleSheet);
els.loadDemoButton.addEventListener("click", loadDemo);
els.sheetSelect.addEventListener("change", (event) => loadWorkbookSheet(event.target.value));
[els.metricSelect, els.yearSelect, els.regionSelect, els.countySelect, els.includePendingCheckbox]
  .forEach((el) => el.addEventListener("change", applyFiltersAndRender));
els.showOwnerCheckbox.addEventListener("change", renderTable);
els.downloadCsvButton.addEventListener("click", downloadFilteredCsv);

document.querySelectorAll(".download-chart").forEach((button) => {
  button.addEventListener("click", () => {
    Plotly.downloadImage(button.dataset.chart, {
      format: "png",
      filename: `pasteurella_${button.dataset.chart}`,
      width: 1400,
      height: 800,
      scale: 2,
    });
  });
});

["annualChart", "countyChart", "classificationChart", "qualityChart"].forEach((id) =>
  renderEmptyChart(id, "載入資料後顯示")
);
