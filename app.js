"use strict";

const REQUIRED_PRICE_HEADERS = ["物料编码", "物料名称", "规格", "单价"];
const REQUIRED_MASTER_HEADERS = ["物料编码", "物料名称", "规格", "物料业务状态"];
const OUTPUT_HEADERS = ["物料编码", "物料名称", "规格", "单价", "物料业务状态"];
const MAX_SHARED_STRINGS_BYTES = 64 * 1024 * 1024;

const state = {
  priceFile: null,
  masterFile: null,
  mergedRows: [],
  processing: false,
};

const els = {
  priceFile: document.getElementById("priceFile"),
  masterFile: document.getElementById("masterFile"),
  priceDrop: document.getElementById("priceDrop"),
  masterDrop: document.getElementById("masterDrop"),
  priceState: document.getElementById("priceState"),
  masterState: document.getElementById("masterState"),
  mergeBtn: document.getElementById("mergeBtn"),
  resetBtn: document.getElementById("resetBtn"),
  progressSection: document.getElementById("progressSection"),
  progressTitle: document.getElementById("progressTitle"),
  progressPercent: document.getElementById("progressPercent"),
  progressBar: document.getElementById("progressBar"),
  progressDetail: document.getElementById("progressDetail"),
  resultSection: document.getElementById("resultSection"),
  resultSummary: document.getElementById("resultSummary"),
  exportBtn: document.getElementById("exportBtn"),
  exportFeedback: document.getElementById("exportFeedback"),
  totalCount: document.getElementById("totalCount"),
  listedCount: document.getElementById("listedCount"),
  retiredCount: document.getElementById("retiredCount"),
  issueCount: document.getElementById("issueCount"),
  warningPanel: document.getElementById("warningPanel"),
  warningText: document.getElementById("warningText"),
  previewBody: document.getElementById("previewBody"),
  errorSection: document.getElementById("errorSection"),
  errorText: document.getElementById("errorText"),
};

class ZipReader {
  constructor(file) {
    this.file = file;
    this.entries = new Map();
  }

  async init() {
    const tailSize = Math.min(this.file.size, 66_000);
    const tailStart = this.file.size - tailSize;
    const tail = new Uint8Array(await this.file.slice(tailStart).arrayBuffer());
    let eocd = -1;
    for (let i = tail.length - 22; i >= 0; i -= 1) {
      if (readUint32(tail, i) === 0x06054b50) {
        eocd = i;
        break;
      }
    }
    if (eocd < 0) throw new Error("文件不是有效的 .xlsx，或文件已经损坏。");

    const centralSize = readUint32(tail, eocd + 12);
    const centralOffset = readUint32(tail, eocd + 16);
    if (centralSize === 0xffffffff || centralOffset === 0xffffffff) {
      throw new Error("暂不支持 Zip64 格式的 Excel 文件。");
    }

    const central = new Uint8Array(
      await this.file.slice(centralOffset, centralOffset + centralSize).arrayBuffer(),
    );
    const decoder = new TextDecoder("utf-8");
    let offset = 0;
    while (offset + 46 <= central.length && readUint32(central, offset) === 0x02014b50) {
      const compression = readUint16(central, offset + 10);
      const compressedSize = readUint32(central, offset + 20);
      const uncompressedSize = readUint32(central, offset + 24);
      const nameLength = readUint16(central, offset + 28);
      const extraLength = readUint16(central, offset + 30);
      const commentLength = readUint16(central, offset + 32);
      const localOffset = readUint32(central, offset + 42);
      const name = decoder.decode(central.slice(offset + 46, offset + 46 + nameLength));
      this.entries.set(normalizeZipPath(name), {
        name: normalizeZipPath(name),
        compression,
        compressedSize,
        uncompressedSize,
        localOffset,
      });
      offset += 46 + nameLength + extraLength + commentLength;
    }
    if (!this.entries.size) throw new Error("Excel 压缩目录为空，无法读取文件。");
    return this;
  }

  has(name) {
    return this.entries.has(normalizeZipPath(name));
  }

  getEntry(name) {
    const entry = this.entries.get(normalizeZipPath(name));
    if (!entry) throw new Error(`Excel 内缺少必要内容：${name}`);
    return entry;
  }

  async stream(name) {
    const entry = this.getEntry(name);
    const header = new Uint8Array(
      await this.file.slice(entry.localOffset, entry.localOffset + 30).arrayBuffer(),
    );
    if (readUint32(header, 0) !== 0x04034b50) throw new Error(`Excel 内容损坏：${name}`);
    const nameLength = readUint16(header, 26);
    const extraLength = readUint16(header, 28);
    const start = entry.localOffset + 30 + nameLength + extraLength;
    const compressed = this.file.slice(start, start + entry.compressedSize).stream();
    if (entry.compression === 0) return compressed;
    if (entry.compression === 8) {
      if (!("DecompressionStream" in window)) {
        throw new Error("浏览器版本过旧，无法分段读取大文件。请改用最新版 Chrome、Edge 或 Safari。");
      }
      return compressed.pipeThrough(new DecompressionStream("deflate-raw"));
    }
    throw new Error(`Excel 使用了暂不支持的压缩方式（${entry.compression}）。`);
  }

  async text(name) {
    return new Response(await this.stream(name)).text();
  }
}

function readUint16(bytes, offset) {
  return bytes[offset] | (bytes[offset + 1] << 8);
}

function readUint32(bytes, offset) {
  return (
    bytes[offset] |
    (bytes[offset + 1] << 8) |
    (bytes[offset + 2] << 16) |
    (bytes[offset + 3] << 24)
  ) >>> 0;
}

function normalizeZipPath(path) {
  const parts = String(path).replace(/^\/+/, "").split("/");
  const clean = [];
  for (const part of parts) {
    if (!part || part === ".") continue;
    if (part === "..") clean.pop();
    else clean.push(part);
  }
  return clean.join("/");
}

function normalize(value) {
  return value == null ? "" : String(value).replace(/\s+/g, " ").trim();
}

function columnFromRef(ref) {
  return (String(ref).match(/^[A-Z]+/) || [""])[0];
}

function parseXml(text, label) {
  const doc = new DOMParser().parseFromString(text, "application/xml");
  const error = doc.querySelector("parsererror");
  if (error) throw new Error(`${label}的 XML 内容无法解析。`);
  return doc;
}

async function getFirstWorksheetPath(zip) {
  if (!zip.has("xl/workbook.xml") || !zip.has("xl/_rels/workbook.xml.rels")) {
    return "xl/worksheets/sheet1.xml";
  }
  const workbookDoc = parseXml(await zip.text("xl/workbook.xml"), "工作簿");
  const sheet = workbookDoc.getElementsByTagName("sheet")[0];
  if (!sheet) throw new Error("Excel 中没有可读取的工作表。");
  const relId = sheet.getAttribute("r:id") || sheet.getAttributeNS("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "id");
  const relsDoc = parseXml(await zip.text("xl/_rels/workbook.xml.rels"), "工作簿关系");
  const rel = Array.from(relsDoc.getElementsByTagName("Relationship")).find(
    (item) => item.getAttribute("Id") === relId,
  );
  if (!rel) return "xl/worksheets/sheet1.xml";
  const target = rel.getAttribute("Target") || "worksheets/sheet1.xml";
  return target.startsWith("/")
    ? normalizeZipPath(target)
    : normalizeZipPath(`xl/${target}`);
}

async function readSharedStrings(zip) {
  if (!zip.has("xl/sharedStrings.xml")) return [];
  const entry = zip.getEntry("xl/sharedStrings.xml");
  if (entry.uncompressedSize > MAX_SHARED_STRINGS_BYTES) {
    throw new Error("共享文本表过大，当前版本无法安全读取。请将源文件另存为新的 .xlsx 后再试。");
  }
  const doc = parseXml(await zip.text("xl/sharedStrings.xml"), "共享文本");
  return Array.from(doc.getElementsByTagName("si"), (item) =>
    Array.from(item.getElementsByTagName("t"), (textNode) => textNode.textContent || "").join(""),
  );
}

function readDomCell(cell, sharedStrings) {
  const type = cell.getAttribute("t") || "";
  if (type === "inlineStr") {
    return Array.from(cell.getElementsByTagName("t"), (node) => node.textContent || "").join("");
  }
  const value = cell.getElementsByTagName("v")[0]?.textContent ?? "";
  if (type === "s") return sharedStrings[Number(value)] ?? "";
  if (type === "str") return value;
  if (value === "") return "";
  const number = Number(value);
  return Number.isFinite(number) ? number : value;
}

function readDomRow(row, sharedStrings) {
  const values = new Map();
  for (const cell of Array.from(row.getElementsByTagName("c"))) {
    values.set(columnFromRef(cell.getAttribute("r")), readDomCell(cell, sharedStrings));
  }
  return values;
}

function mapHeaderColumns(values, requiredHeaders, fileLabel) {
  const byName = new Map();
  for (const [column, value] of values.entries()) byName.set(normalize(value), column);
  const missing = requiredHeaders.filter((header) => !byName.has(header));
  if (missing.length) {
    throw new Error(`${fileLabel}缺少列：${missing.join("、")}。请确认选中了正确文件。`);
  }
  return Object.fromEntries(requiredHeaders.map((header) => [header, byName.get(header)]));
}

async function readPriceWorkbook(file) {
  updateProgress(5, "读取价格列表", "正在检查工作表和共享文本");
  const zip = await new ZipReader(file).init();
  const sheetPath = await getFirstWorksheetPath(zip);
  const sharedStrings = await readSharedStrings(zip);
  const doc = parseXml(await zip.text(sheetPath), "价格列表工作表");
  const rows = Array.from(doc.getElementsByTagName("row"));
  if (!rows.length) throw new Error("价格列表查询中没有数据。");

  const headerColumns = mapHeaderColumns(
    readDomRow(rows[0], sharedStrings),
    REQUIRED_PRICE_HEADERS,
    "价格列表查询",
  );
  const result = [];
  const duplicates = new Set();
  const seen = new Set();
  for (let index = 1; index < rows.length; index += 1) {
    const values = readDomRow(rows[index], sharedStrings);
    const code = normalize(values.get(headerColumns["物料编码"]));
    if (!code) continue;
    if (seen.has(code)) duplicates.add(code);
    seen.add(code);
    const rawPrice = values.get(headerColumns["单价"]);
    const numberPrice = typeof rawPrice === "number" ? rawPrice : Number(rawPrice);
    result.push({
      code,
      name: normalize(values.get(headerColumns["物料名称"])),
      spec: normalize(values.get(headerColumns["规格"])),
      price: Number.isFinite(numberPrice) ? numberPrice : normalize(rawPrice),
    });
  }
  if (!result.length) throw new Error("价格列表查询没有可用的物料记录。");
  return { rows: result, duplicates: Array.from(duplicates) };
}

function decodeXmlText(text) {
  return String(text)
    .replace(/&#x([0-9a-f]+);/gi, (_, hex) => String.fromCodePoint(parseInt(hex, 16)))
    .replace(/&#(\d+);/g, (_, decimal) => String.fromCodePoint(Number(decimal)))
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, "&");
}

function readRegexCell(body, type, sharedStrings) {
  if (type === "inlineStr") {
    let text = "";
    const textPattern = /<t(?:\s[^>]*)?>([\s\S]*?)<\/t>/g;
    let textMatch;
    while ((textMatch = textPattern.exec(body))) text += decodeXmlText(textMatch[1]);
    return text;
  }
  const value = body.match(/<v(?:\s[^>]*)?>([\s\S]*?)<\/v>/)?.[1] ?? "";
  const decoded = decodeXmlText(value);
  if (type === "s") return sharedStrings[Number(decoded)] ?? "";
  return decoded;
}

function parseRegexRow(rowXml, sharedStrings, wantedColumns = null) {
  const values = new Map();
  const cellPattern = /<c\b([^>]*)>([\s\S]*?)<\/c>/g;
  let match;
  while ((match = cellPattern.exec(rowXml))) {
    const attributes = match[1];
    const ref = attributes.match(/\br="([^"]+)"/)?.[1] || "";
    const column = columnFromRef(ref);
    if (!column || (wantedColumns && !wantedColumns.has(column))) continue;
    const type = attributes.match(/\bt="([^"]+)"/)?.[1] || "";
    values.set(column, readRegexCell(match[2], type, sharedStrings));
  }
  return values;
}

async function readMasterWorkbook(file, priceCodes) {
  updateProgress(18, "读取物料主数据", "正在建立分段读取通道");
  const zip = await new ZipReader(file).init();
  const sheetPath = await getFirstWorksheetPath(zip);
  const sheetEntry = zip.getEntry(sheetPath);
  const sharedStrings = await readSharedStrings(zip);
  const stream = await zip.stream(sheetPath);
  const reader = stream.getReader();
  const decoder = new TextDecoder("utf-8");
  const matches = new Map();
  const conflicts = [];
  let buffer = "";
  let bytesRead = 0;
  let masterRowCount = 0;
  let headerColumns = null;
  let wantedColumns = null;
  let lastShownPercent = -1;

  const consumeRows = () => {
    while (true) {
      const start = buffer.indexOf("<row");
      if (start < 0) {
        buffer = buffer.slice(-8);
        return;
      }
      const end = buffer.indexOf("</row>", start);
      if (end < 0) {
        if (start > 0) buffer = buffer.slice(start);
        return;
      }
      const rowXml = buffer.slice(start, end + 6);
      buffer = buffer.slice(end + 6);
      masterRowCount += 1;

      if (!headerColumns) {
        const headerValues = parseRegexRow(rowXml, sharedStrings);
        headerColumns = mapHeaderColumns(headerValues, REQUIRED_MASTER_HEADERS, "物料主数据");
        wantedColumns = new Set(Object.values(headerColumns));
        continue;
      }

      const values = parseRegexRow(rowXml, sharedStrings, wantedColumns);
      const code = normalize(values.get(headerColumns["物料编码"]));
      if (!code || !priceCodes.has(code)) continue;
      const record = {
        name: normalize(values.get(headerColumns["物料名称"])),
        spec: normalize(values.get(headerColumns["规格"])),
        status: normalize(values.get(headerColumns["物料业务状态"])),
      };
      const existing = matches.get(code);
      if (!existing || (!existing.status && record.status)) {
        matches.set(code, record);
      } else if (
        existing.name !== record.name ||
        existing.spec !== record.spec ||
        (existing.status && record.status && existing.status !== record.status)
      ) {
        conflicts.push(code);
      }
    }
  };

  while (true) {
    const { value, done } = await reader.read();
    if (done) break;
    bytesRead += value.byteLength;
    buffer += decoder.decode(value, { stream: true });
    consumeRows();
    const innerPercent = Math.min(100, Math.floor((bytesRead / sheetEntry.uncompressedSize) * 100));
    if (innerPercent >= lastShownPercent + 2) {
      lastShownPercent = innerPercent;
      updateProgress(
        20 + Math.round(innerPercent * 0.62),
        "读取物料主数据",
        `已扫描 ${masterRowCount.toLocaleString("zh-CN")} 行，匹配到 ${matches.size.toLocaleString("zh-CN")} 个价格物料`,
      );
      await new Promise((resolve) => setTimeout(resolve, 0));
    }
  }
  buffer += decoder.decode();
  consumeRows();
  if (!headerColumns) throw new Error("物料主数据中没有可识别的表头。");
  return { matches, conflicts: Array.from(new Set(conflicts)), masterRowCount: Math.max(0, masterRowCount - 1) };
}

function mergeData(priceData, masterData) {
  updateProgress(84, "合并并校验", "正在按物料编码匹配，并核对名称与规格");
  const mergedRows = [];
  const missing = [];
  const mismatches = [];
  const statusCounts = new Map();

  for (const price of priceData.rows) {
    const master = masterData.matches.get(price.code);
    const status = master ? master.status || "状态空白" : "未匹配";
    if (!master) missing.push(price.code);
    else if (price.name !== master.name || price.spec !== master.spec) mismatches.push(price.code);
    statusCounts.set(status, (statusCounts.get(status) || 0) + 1);
    mergedRows.push([price.code, price.name, price.spec, price.price, status]);
  }
  return {
    mergedRows,
    missing,
    mismatches,
    statusCounts,
    duplicatePriceCodes: priceData.duplicates,
    conflictingMasterCodes: masterData.conflicts,
    masterRowCount: masterData.masterRowCount,
  };
}

async function startMerge() {
  if (state.processing || !state.priceFile || !state.masterFile) return;
  state.processing = true;
  state.mergedRows = [];
  els.mergeBtn.disabled = true;
  els.resetBtn.disabled = true;
  els.resultSection.classList.add("hidden");
  els.errorSection.classList.add("hidden");
  els.progressSection.classList.remove("hidden");
  updateProgress(1, "正在准备文件", "所有数据均在本机浏览器内处理，不会上传");

  try {
    const priceData = await readPriceWorkbook(state.priceFile);
    updateProgress(15, "价格列表已读取", `共 ${priceData.rows.length.toLocaleString("zh-CN")} 条物料，准备读取主数据`);
    const priceCodes = new Set(priceData.rows.map((row) => row.code));
    const masterData = await readMasterWorkbook(state.masterFile, priceCodes);
    const result = mergeData(priceData, masterData);
    state.mergedRows = result.mergedRows;
    updateProgress(100, "合并完成", "结果已生成，可以导出 Excel");
    showResult(result);
  } catch (error) {
    console.error(error);
    showError(error instanceof Error ? error.message : String(error));
  } finally {
    state.processing = false;
    els.resetBtn.disabled = false;
    els.mergeBtn.disabled = !(state.priceFile && state.masterFile);
  }
}

function updateProgress(percent, title, detail) {
  const safePercent = Math.max(0, Math.min(100, Number(percent) || 0));
  els.progressTitle.textContent = title;
  els.progressDetail.textContent = detail;
  els.progressPercent.textContent = `${Math.round(safePercent)}%`;
  els.progressBar.style.width = `${safePercent}%`;
}

function showResult(result) {
  const listed = result.statusCounts.get("已上市") || 0;
  const retired = result.statusCounts.get("已淘汰") || 0;
  const issueTotal =
    result.missing.length +
    result.mismatches.length +
    result.duplicatePriceCodes.length +
    result.conflictingMasterCodes.length;
  els.totalCount.textContent = result.mergedRows.length.toLocaleString("zh-CN");
  els.listedCount.textContent = listed.toLocaleString("zh-CN");
  els.retiredCount.textContent = retired.toLocaleString("zh-CN");
  els.issueCount.textContent = issueTotal.toLocaleString("zh-CN");
  els.resultSummary.textContent = `已扫描物料主数据 ${result.masterRowCount.toLocaleString("zh-CN")} 行，输出 ${result.mergedRows.length.toLocaleString("zh-CN")} 条价格记录。`;

  const warnings = [];
  if (result.missing.length) warnings.push(`${result.missing.length} 个编码未在物料主数据中找到（导出时标记为“未匹配”）`);
  if (result.mismatches.length) warnings.push(`${result.mismatches.length} 个编码的名称或规格不一致（仍按物料编码匹配）`);
  if (result.duplicatePriceCodes.length) warnings.push(`${result.duplicatePriceCodes.length} 个物料编码在价格表中重复`);
  if (result.conflictingMasterCodes.length) warnings.push(`${result.conflictingMasterCodes.length} 个物料编码在主数据中存在冲突记录`);
  els.warningPanel.classList.toggle("hidden", warnings.length === 0);
  els.warningText.textContent = warnings.join("；");

  renderPreview(result.mergedRows.slice(0, 12));
  els.resultSection.classList.remove("hidden");
  setTimeout(() => els.resultSection.scrollIntoView({ behavior: "smooth", block: "start" }), 120);
}

function renderPreview(rows) {
  els.previewBody.textContent = "";
  for (const row of rows) {
    const tr = document.createElement("tr");
    row.forEach((value, index) => {
      const td = document.createElement("td");
      if (index === 3) {
        td.className = "number-cell";
        td.textContent = typeof value === "number" ? value.toFixed(2) : String(value ?? "");
      } else if (index === 4) {
        const badge = document.createElement("span");
        badge.className = `status-badge ${
          value === "已上市" ? "status-listed" : value === "已淘汰" ? "status-retired" : "status-missing"
        }`;
        badge.textContent = String(value ?? "");
        td.appendChild(badge);
      } else {
        td.textContent = String(value ?? "");
        td.title = String(value ?? "");
      }
      tr.appendChild(td);
    });
    els.previewBody.appendChild(tr);
  }
}

function showError(message) {
  els.progressSection.classList.add("hidden");
  els.resultSection.classList.add("hidden");
  els.errorText.textContent = message;
  els.errorSection.classList.remove("hidden");
  els.errorSection.scrollIntoView({ behavior: "smooth", block: "center" });
}

function exportWorkbook() {
  window.__exportDiagnostics = {
    clicked: true,
    rows: state.mergedRows.length,
    xlsxLoaded: typeof XLSX !== "undefined",
  };
  if (!state.mergedRows.length) {
    window.__exportDiagnostics.error = "没有可导出的合并结果";
    return;
  }
  if (typeof XLSX === "undefined") {
    showError("导出组件未正确加载。请确认 vendor 文件夹与工具页面放在同一目录后重试。");
    return;
  }
  try {
    const workbook = XLSX.utils.book_new();
    const worksheet = XLSX.utils.aoa_to_sheet([OUTPUT_HEADERS, ...state.mergedRows]);
    worksheet["!cols"] = [{ wch: 15 }, { wch: 62 }, { wch: 24 }, { wch: 13 }, { wch: 16 }];
    worksheet["!autofilter"] = { ref: `A1:E${state.mergedRows.length + 1}` };
    for (let row = 2; row <= state.mergedRows.length + 1; row += 1) {
      const codeCell = worksheet[`A${row}`];
      if (codeCell) {
        codeCell.t = "s";
        codeCell.v = String(codeCell.v ?? "");
        codeCell.z = "@";
      }
      const priceCell = worksheet[`D${row}`];
      if (priceCell && priceCell.t === "n") priceCell.z = "#,##0.00";
    }
    XLSX.utils.book_append_sheet(workbook, worksheet, "上市产品价格查询");
    workbook.Props = {
      Title: "上市产品价格查询",
      Subject: "价格列表与物料主数据合并结果",
      Author: "上市产品价格合并工具",
      CreatedDate: new Date(),
    };
    const today = new Date();
    const date = `${today.getFullYear()}${String(today.getMonth() + 1).padStart(2, "0")}${String(today.getDate()).padStart(2, "0")}`;
    const outputBytes = XLSX.write(workbook, { bookType: "xlsx", type: "array", compression: true });
    const outputBlob = new Blob([outputBytes], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });
    window.__lastExportInfo = {
      name: `上市产品价格查询_${date}.xlsx`,
      size: outputBlob.size,
      signature: Array.from(new Uint8Array(outputBytes).slice(0, 4)),
    };
    window.__exportDiagnostics.output = window.__lastExportInfo;
    els.exportBtn.dataset.exportSize = String(outputBlob.size);
    els.exportBtn.dataset.exportSignature = window.__lastExportInfo.signature.join(",");
    const downloadUrl = URL.createObjectURL(outputBlob);
    const downloadLink = document.createElement("a");
    downloadLink.href = downloadUrl;
    downloadLink.download = window.__lastExportInfo.name;
    downloadLink.style.display = "none";
    document.body.appendChild(downloadLink);
    downloadLink.click();
    downloadLink.remove();
    setTimeout(() => URL.revokeObjectURL(downloadUrl), 30_000);
    els.exportFeedback.textContent = `已生成 ${window.__lastExportInfo.name}`;
    els.exportFeedback.classList.remove("hidden");
  } catch (error) {
    window.__exportDiagnostics.error = error instanceof Error ? error.message : String(error);
    els.exportBtn.dataset.exportError = window.__exportDiagnostics.error;
    showError(`导出失败：${window.__exportDiagnostics.error}`);
  }
}

function formatFileSize(bytes) {
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(0)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}

function setFile(kind, file) {
  if (!file) return;
  if (!file.name.toLowerCase().endsWith(".xlsx")) {
    showError("目前仅支持 .xlsx 文件，请先将源文件另存为 .xlsx 后再导入。");
    return;
  }
  state[`${kind}File`] = file;
  const card = kind === "price" ? els.priceDrop : els.masterDrop;
  const label = kind === "price" ? els.priceState : els.masterState;
  card.classList.add("selected");
  label.textContent = `${file.name} · ${formatFileSize(file.size)}`;
  els.errorSection.classList.add("hidden");
  els.resetBtn.disabled = false;
  els.mergeBtn.disabled = !(state.priceFile && state.masterFile);
  if (state.priceFile && state.masterFile && !state.processing) setTimeout(startMerge, 180);
}

function bindFileCard(kind, card, input) {
  card.addEventListener("click", () => !state.processing && input.click());
  card.addEventListener("keydown", (event) => {
    if ((event.key === "Enter" || event.key === " ") && !state.processing) {
      event.preventDefault();
      input.click();
    }
  });
  input.addEventListener("change", () => setFile(kind, input.files?.[0]));
  for (const eventName of ["dragenter", "dragover"]) {
    card.addEventListener(eventName, (event) => {
      event.preventDefault();
      if (!state.processing) card.classList.add("dragging");
    });
  }
  for (const eventName of ["dragleave", "drop"]) {
    card.addEventListener(eventName, (event) => {
      event.preventDefault();
      card.classList.remove("dragging");
    });
  }
  card.addEventListener("drop", (event) => {
    if (!state.processing) setFile(kind, event.dataTransfer?.files?.[0]);
  });
}

function resetAll() {
  if (state.processing) return;
  state.priceFile = null;
  state.masterFile = null;
  state.mergedRows = [];
  els.priceFile.value = "";
  els.masterFile.value = "";
  els.priceDrop.classList.remove("selected");
  els.masterDrop.classList.remove("selected");
  els.priceState.textContent = "点击选择或拖入 .xlsx 文件";
  els.masterState.textContent = "点击选择或拖入 .xlsx 文件";
  els.mergeBtn.disabled = true;
  els.resetBtn.disabled = true;
  els.progressSection.classList.add("hidden");
  els.resultSection.classList.add("hidden");
  els.errorSection.classList.add("hidden");
  els.exportFeedback.classList.add("hidden");
  els.exportFeedback.textContent = "";
}

bindFileCard("price", els.priceDrop, els.priceFile);
bindFileCard("master", els.masterDrop, els.masterFile);
els.mergeBtn.addEventListener("click", startMerge);
els.exportBtn.addEventListener("click", exportWorkbook);
els.resetBtn.addEventListener("click", resetAll);

if (typeof XLSX === "undefined") {
  showError("导出组件未加载。请保留 vendor 文件夹，并从工具目录中打开本页面。");
}
