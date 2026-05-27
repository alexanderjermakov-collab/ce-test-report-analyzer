const APP_NAME = "CE Test Report Analyzer";
const APP_VERSION = "1.0";
const APP_AUTHOR = "Alex Jermakov";

const state = {
  pdfFiles: [],
  cstFile: null,
  summaryBlob: null,
  logBlob: null,
  logs: [],
};

const els = {
  pdfInput: document.querySelector("#pdfInput"),
  cstInput: document.querySelector("#cstInput"),
  pdfDropZone: document.querySelector("#pdfDropZone"),
  cstDropZone: document.querySelector("#cstDropZone"),
  pdfList: document.querySelector("#pdfList"),
  pdfCount: document.querySelector("#pdfCount"),
  cstState: document.querySelector("#cstState"),
  cstFileName: document.querySelector("#cstFileName"),
  asfState: document.querySelector("#asfState"),
  asfInfo: document.querySelector("#asfInfo"),
  logState: document.querySelector("#logState"),
  logInfo: document.querySelector("#logInfo"),
  analyzeButton: document.querySelector("#analyzeButton"),
  downloadSummaryButton: document.querySelector("#downloadSummaryButton"),
  downloadLogButton: document.querySelector("#downloadLogButton"),
  progressBar: document.querySelector("#progressBar"),
  progressText: document.querySelector("#progressText"),
  statusLog: document.querySelector("#statusLog"),
};

const pdfjsLib = await import("https://cdnjs.cloudflare.com/ajax/libs/pdf.js/4.10.38/pdf.min.mjs");
pdfjsLib.GlobalWorkerOptions.workerSrc = "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/4.10.38/pdf.worker.min.mjs";

function log(message, level = "INFO") {
  const line = `${new Date().toISOString()} [${level}] ${message}`;
  state.logs.push(line);
  els.statusLog.textContent = state.logs.slice(-120).join("\n");
  els.statusLog.scrollTop = els.statusLog.scrollHeight;
}

function setProgress(value, text) {
  els.progressBar.value = value;
  els.progressText.textContent = text;
}

function formatBytes(bytes) {
  if (!bytes) return "0 B";
  const units = ["B", "KB", "MB", "GB"];
  const index = Math.min(Math.floor(Math.log(bytes) / Math.log(1024)), units.length - 1);
  return `${(bytes / 1024 ** index).toFixed(index ? 1 : 0)} ${units[index]}`;
}

function renderFileState() {
  els.pdfList.replaceChildren();
  state.pdfFiles.forEach((file) => {
    const li = document.createElement("li");
    li.textContent = `${file.name} (${formatBytes(file.size)})`;
    els.pdfList.append(li);
  });
  if (!state.pdfFiles.length) {
    const li = document.createElement("li");
    li.textContent = "No PDF files loaded";
    els.pdfList.append(li);
  }
  els.pdfCount.textContent = `${state.pdfFiles.length} file${state.pdfFiles.length === 1 ? "" : "s"}`;

  els.cstState.textContent = state.cstFile ? "Loaded" : "Not selected";
  els.cstFileName.textContent = state.cstFile ? `${state.cstFile.name} (${formatBytes(state.cstFile.size)})` : "No workbook loaded";
}

function resetOutputs() {
  state.summaryBlob = null;
  state.logBlob = null;
  els.downloadSummaryButton.disabled = true;
  els.downloadLogButton.disabled = true;
  els.asfState.textContent = "Waiting";
  els.asfInfo.textContent = "Run analysis to generate XLSX output.";
  els.logState.textContent = "Ready";
  els.logInfo.textContent = "Technical log will be generated during analysis.";
}

function acceptPdfFiles(files) {
  const incoming = [...files].filter((file) => file.type === "application/pdf" || file.name.toLowerCase().endsWith(".pdf"));
  const keyed = new Map(state.pdfFiles.map((file) => [`${file.name}:${file.size}`, file]));
  incoming.forEach((file) => keyed.set(`${file.name}:${file.size}`, file));
  state.pdfFiles = [...keyed.values()];
  resetOutputs();
  renderFileState();
  log(`Loaded ${incoming.length} CETR PDF file(s).`);
}

function acceptCstFile(files) {
  const file = [...files].find((item) => /\.(xlsx|xls)$/i.test(item.name));
  if (!file) {
    log("No XLSX/XLS CST workbook was selected.", "WARN");
    return;
  }
  state.cstFile = file;
  resetOutputs();
  renderFileState();
  log(`Loaded CST workbook: ${file.name}.`);
}

function setupDropZone(zone, onFiles) {
  zone.addEventListener("dragover", (event) => {
    event.preventDefault();
    zone.classList.add("drag-over");
  });
  zone.addEventListener("dragleave", () => zone.classList.remove("drag-over"));
  zone.addEventListener("drop", (event) => {
    event.preventDefault();
    zone.classList.remove("drag-over");
    onFiles(event.dataTransfer.files);
  });
}

els.pdfInput.addEventListener("change", (event) => acceptPdfFiles(event.target.files));
els.cstInput.addEventListener("change", (event) => acceptCstFile(event.target.files));
setupDropZone(els.pdfDropZone, acceptPdfFiles);
setupDropZone(els.cstDropZone, acceptCstFile);

async function extractPdfText(file) {
  const bytes = await file.arrayBuffer();
  const pdf = await pdfjsLib.getDocument({ data: bytes }).promise;
  const pages = [];
  for (let pageNumber = 1; pageNumber <= pdf.numPages; pageNumber += 1) {
    const page = await pdf.getPage(pageNumber);
    const textContent = await page.getTextContent();
    pages.push(textContent.items.map((item) => item.str).join(" "));
  }
  return pages.join("\n");
}

function cleanText(text) {
  return text
    .replace(/\u0000/g, "")
    .replace(/([124])T\s*-\s*C/g, "$1T-C")
    .replace(/PR\s*([0-9])/g, "PR$1")
    .replace(/LT\s*-\s*([0-9])/g, "LT-$1")
    .replace(/\s+/g, " ")
    .trim();
}

function matchFirst(text, patterns, fallback = "") {
  for (const pattern of patterns) {
    const match = text.match(pattern);
    if (match?.[1]) return cleanText(match[1]);
  }
  return fallback;
}

function extractListNearLabel(text, labels) {
  for (const label of labels) {
    const stop = [
      "Product Description", "Brand Name", "Brand names", "Applicant", "Sample Receipt",
      "Test Conducted Date", "Date Received", "Electrical Rating", "Serial", "Support Equipment",
      "Specification", "Verification Issuing", "Test Result", "Conclusion",
    ].join("|");
    const pattern = new RegExp(`${label}\\s*\\(?s?\\)?\\s*[:\\-]?\\s*([A-Z0-9][\\s\\S]{2,620}?)(?=\\s(?:${stop})\\b|$)`, "i");
    const match = text.match(pattern);
    if (match?.[1]) return splitNames(match[1]);
  }
  return [];
}

function splitNames(value) {
  return [...new Set(
    value
      .split(/[,;\n]|(?:\s+\/\s+)|(?:\s{2,})|(?:\s+and\s+)/i)
      .map((item) => cleanText(item))
      .map((item) => item.replace(/\b(model|models|brand|brands|name|names)\b/gi, "").trim())
      .map((item) => item.replace(/\s*\/\s*$/, "").trim())
      .filter((item) => !item.includes("*"))
      .filter((item) => !/\b(can from|blank|not labelled|n\/a)\b/i.test(item))
      .filter((item) => !/^(not kick off|to remove|tbd|comments?)$/i.test(item))
      .filter((item) => item.length >= 2 && item.length <= 80)
  )];
}

function inferModels(text) {
  const labelled = extractListNearLabel(text, [
    "Models/Type References", "Model/Type reference", "Type Number", "Type Number(s)",
    "Model No.", "Model", "Models", "Sample Description",
  ]);
  const candidates = text.match(/\b(?:[124]T-C|PR|M[HU]DV|MSDV|LD|F|LT-?)[A-Z0-9xX*_-]{5,32}\b/g) || [];
  const normalized = candidates
    .map((item) => cleanText(item))
    .map((item) => item.replace(/[.,;:)]$/, ""))
    .filter((item) => !item.includes("*"))
    .filter((item) => !/^(20[0-9]{2}|19[0-9]{2})$/.test(item))
    .filter((item) => !/^\d+(?:V|HZ|W|MM|CM|KG)$/i.test(item));
  return [...new Set([...labelled, ...normalized])].slice(0, 80);
}

function inferBrands(text) {
  const labelled = extractListNearLabel(text, ["Brand Name", "Brand Name(s)", "Brand names", "Trade Mark", "Trade Mark(s)", "Brand"]);
  const knownBrands = ["AMTC", "Sharp", "Polaroid", "Toshiba", "Hitachi", "JVC", "Blaupunkt", "Daewoo", "Telefunken", "Nokia", "Logik", "Vitado"];
  const found = knownBrands.filter((brand) => new RegExp(`\\b${brand}\\b`, "i").test(text));
  return [...new Set([...labelled, ...found].map((brand) => brand.replace(/Polariod/i, "Polaroid")))];
}

function inferStandards(text) {
  const standards = text.match(/\b(?:EN|IEC|ETSI EN|EN IEC|BS EN)\s*[0-9]{3,6}(?:[- ][0-9A-Z]+)*(?::\s*[0-9]{4})?(?:\+A[0-9]+(?::[0-9]{4})?)?/gi) || [];
  return [...new Set(standards.map((standard) => cleanText(standard).replace(/\s+/g, " ")))].slice(0, 20).join("; ");
}

function analyzeReport(file, rawText) {
  const text = cleanText(rawText);
  const conclusion = /\bfail(?:ed|ure)?\b/i.test(text)
    ? "FAIL"
    : /\b(pass(?:ed)?|complied|met the requirement|does meet the requirement)\b/i.test(text)
      ? "PASS"
      : "";
  return {
    fileName: file.name,
    testingDate: matchFirst(text, [
      /(?:test conducted date|test date\(s\)|test date|date test conducted|testing date|date of test|test period)\s*[:\-]?\s*([0-9]{1,2}\s+[A-Za-z]{3,9}\s+[0-9]{4}(?:\s*(?:to|-|,)\s*[0-9]{1,2}\s+[A-Za-z]{3,9}\s+[0-9]{4})*)/i,
      /(?:date received|date of issue|issue date)\s*[:\-]?\s*([A-Za-z]{3,9}\.?\s+[0-9]{1,2},\s+[0-9]{4})/i,
      /(?:test(?:ing)? date|date of test|test period)\s*[:\-]?\s*([0-9]{1,2}[./-][0-9]{1,2}[./-][0-9]{2,4}(?:\s*(?:to|-)\s*[0-9]{1,2}[./-][0-9]{1,2}[./-][0-9]{2,4})?)/i,
      /(?:test(?:ing)? date|date of test|test period)\s*[:\-]?\s*([0-9]{4}[./-][0-9]{1,2}[./-][0-9]{1,2})/i,
    ]),
    laboratory: matchFirst(text, [
      /(?:name of testing laboratory preparing the report|test site and location|verification issuing office name & address)\s*[:\-]?\s*([^:]{3,180}?)(?=\s(?:applicant|sample|test result|test specification|address|product|report))/i,
      /(?:test laboratory|testing laboratory|laboratory|test lab)\s*[:\-]?\s*([^:]{3,120}?)(?=\s(?:applicant|manufacturer|report|standard|test))/i,
    ]),
    applicant: matchFirst(text, [
      /(?:applicant(?:’s name|'s name)?|applicant name & address|client)\s*[:\-]?\s*([^:]{3,140}?)(?=\s(?:MTC Industry|address|manufacturer|sample|test|standard|model|report|product))/i,
    ]),
    standard: inferStandards(text) || matchFirst(text, [
      /(?:standard|test standard|used standard)\s*[:\-]?\s*((?:EN|IEC|BS EN|ETSI)\s*[0-9][A-Z0-9:./\- ]{3,80})/i,
      /\b((?:EN|IEC|BS EN)\s*[0-9]{4,6}(?:-[0-9])?(?::[0-9]{4})?)\b/i,
    ]),
    conclusion,
    models: inferModels(text),
    brands: inferBrands(text),
    rawTextLength: rawText.length,
  };
}

async function readWorkbook(file) {
  const bytes = await file.arrayBuffer();
  return XLSX.read(bytes, { type: "array", cellDates: true });
}

function sheetRows(workbook, sheetName) {
  return XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1, defval: "" });
}

function normalizeLabel(value) {
  return String(value || "")
    .toLowerCase()
    .replace(/\s+/g, " ")
    .trim();
}

function rowLabelMatches(label, aliases) {
  const normalized = normalizeLabel(label);
  return aliases.some((alias) => normalized.includes(alias));
}

function findTransposedLabelRows(rows) {
  const labels = {
    factoryModel: ["amtc factory model name", "factory model", "vendor model name"],
    brand: ["sce tv brand", "tv brand", "brand"],
    model: ["sce model name", "model name"],
    ftp: ["ftp storage folder", "ftp storage", "ftp"],
    project: ["tv project name", "project name", "certification project year"],
  };
  const indexes = {};
  rows.forEach((row, rowIndex) => {
    const label = row[0];
    Object.entries(labels).forEach(([field, aliases]) => {
      if (indexes[field] === undefined && rowLabelMatches(label, aliases)) {
        indexes[field] = rowIndex;
      }
    });
  });
  return indexes;
}

function rowValue(rows, rowIndex, columnIndex) {
  if (rowIndex === undefined) return "";
  return String(rows[rowIndex]?.[columnIndex] ?? "").trim();
}

function sheetLooksTransposed(rows) {
  const indexes = findTransposedLabelRows(rows);
  return indexes.model !== undefined || indexes.factoryModel !== undefined;
}

function transposedSheetRecords(rows, sheetName) {
  const indexes = findTransposedLabelRows(rows);
  const maxColumns = rows.reduce((max, row) => Math.max(max, row.length), 0);
  const records = [];
  for (let columnIndex = 1; columnIndex < maxColumns; columnIndex += 1) {
    const record = {
      sheetName,
      factoryModel: rowValue(rows, indexes.factoryModel, columnIndex),
      brand: rowValue(rows, indexes.brand, columnIndex),
      model: rowValue(rows, indexes.model, columnIndex),
      ftp: rowValue(rows, indexes.ftp, columnIndex),
      project: rowValue(rows, indexes.project, columnIndex),
      matchedRow: "",
    };
    const joined = [record.factoryModel, record.brand, record.model, record.ftp, record.project].filter(Boolean).join(" | ");
    if (!joined) continue;
    record.matchedRow = joined;
    records.push(record);
  }
  return records;
}

function findHeaderIndex(headers, names) {
  const normalizedNames = names.map((name) => name.toLowerCase());
  return headers.findIndex((header) => normalizedNames.some((name) => String(header).toLowerCase().includes(name)));
}

function rowSheetRecords(rows, sheetName) {
  const headerRowIndex = rows.findIndex((row) => row.some((cell) => /model|brand|project|ftp|factory/i.test(String(cell))));
  const headers = headerRowIndex >= 0 ? rows[headerRowIndex] : rows[0];
  const dataRows = rows.slice(Math.max(headerRowIndex + 1, 1));
  const indexes = {
    factoryModel: findHeaderIndex(headers, ["amtc factory model", "factory model", "vendor model", "amtc"]),
    brand: findHeaderIndex(headers, ["sce tv brand", "tv brand", "brand"]),
    model: findHeaderIndex(headers, ["sce model name", "model name", "globalname", "model"]),
    ftp: findHeaderIndex(headers, ["ftp storage folder", "ftp", "storage"]),
    project: findHeaderIndex(headers, ["tv project name", "project"]),
  };
  return dataRows.map((row) => {
    const cells = row.map((cell) => String(cell ?? "").trim());
    return {
      sheetName,
      factoryModel: indexes.factoryModel >= 0 ? cells[indexes.factoryModel] : "",
      brand: indexes.brand >= 0 ? cells[indexes.brand] : "",
      model: indexes.model >= 0 ? cells[indexes.model] : "",
      ftp: indexes.ftp >= 0 ? cells[indexes.ftp] : "",
      project: indexes.project >= 0 ? cells[indexes.project] : "",
      matchedRow: cells.join(" | "),
    };
  });
}

function searchableCstText(record) {
  return [record.factoryModel, record.brand, record.model, record.ftp, record.project]
    .filter(Boolean)
    .join(" ")
    .toLowerCase();
}

function findModelInCst(workbook, modelName) {
  const needle = modelName.toLowerCase();
  const matches = [];
  for (const sheetName of workbook.SheetNames) {
    const rows = sheetRows(workbook, sheetName);
    if (!rows.length) continue;
    const records = sheetLooksTransposed(rows) ? transposedSheetRecords(rows, sheetName) : rowSheetRecords(rows, sheetName);
    records.forEach((record) => {
      if (searchableCstText(record).includes(needle)) matches.push(record);
    });
  }
  return matches;
}

function buildSummaryWorkbook(reports, cstFileName) {
  const executedAt = new Date();
  const wb = XLSX.utils.book_new();
  const titleRows = [
    [APP_NAME],
    ["Application version", APP_VERSION],
    ["Application author", APP_AUTHOR],
    ["Analysis execution date/time", executedAt.toLocaleString()],
    ["Functionality", "Analyzes CE Test Report PDF files and checks TV model synchronization against all worksheets in a CST workbook."],
    ["CETR file names", reports.map((report) => report.fileName).join("; ")],
    ["CST file name", cstFileName],
    [],
  ];

  const summaryRows = reports.map((report) => ({
    "CETR File name": report.fileName,
    "Testing date": report.testingDate,
    "Test Laboratory name": report.laboratory,
    "Applicant company name": report.applicant,
    "Used test standard name": report.standard,
    "Test conclusion": report.conclusion,
    "Tested TV samples model names": report.models.join("; "),
    "Tested TV samples brand names": report.brands.join("; "),
    "CST worksheet names where models were found": [...new Set(report.cstMatches.map((match) => match.sheetName))].join("; "),
    "AMTC factory model name": [...new Set(report.cstMatches.map((match) => match.factoryModel).filter(Boolean))].join("; "),
    "SCE TV Brand": [...new Set(report.cstMatches.map((match) => match.brand).filter(Boolean))].join("; "),
    "SCE Model Name": [...new Set(report.cstMatches.map((match) => match.model).filter(Boolean))].join("; "),
    "FTP storage folder": [...new Set(report.cstMatches.map((match) => match.ftp).filter(Boolean))].join("; "),
    "TV project name": [...new Set(report.cstMatches.map((match) => match.project).filter(Boolean))].join("; "),
    "TV model names missing in CETR": report.modelsMissingInCetr.join("; "),
    "Models not found in CST": report.modelsNotFoundInCst.join("; "),
  }));

  const ws = XLSX.utils.aoa_to_sheet(titleRows);
  XLSX.utils.sheet_add_json(ws, summaryRows, { origin: `A${titleRows.length + 1}` });
  ws["!cols"] = [
    { wch: 34 }, { wch: 18 }, { wch: 28 }, { wch: 28 }, { wch: 24 }, { wch: 14 },
    { wch: 42 }, { wch: 30 }, { wch: 34 }, { wch: 26 }, { wch: 22 }, { wch: 28 },
    { wch: 30 }, { wch: 24 }, { wch: 34 }, { wch: 34 },
  ];
  XLSX.utils.book_append_sheet(wb, ws, "Analysis Summary");

  const detailRows = [];
  reports.forEach((report) => {
    report.cstMatches.forEach((match) => {
      detailRows.push({
        "CETR File name": report.fileName,
        "Matched model from CETR": match.sourceModel,
        "CST worksheet": match.sheetName,
        "AMTC factory model name": match.factoryModel,
        "SCE TV Brand": match.brand,
        "SCE Model Name": match.model,
        "FTP storage folder": match.ftp,
        "TV project name": match.project,
        "Matched CST row": match.matchedRow,
      });
    });
  });
  const detailsWs = XLSX.utils.json_to_sheet(detailRows.length ? detailRows : [{ "Result": "No CST matches found" }]);
  detailsWs["!cols"] = [{ wch: 34 }, { wch: 24 }, { wch: 24 }, { wch: 28 }, { wch: 22 }, { wch: 28 }, { wch: 30 }, { wch: 24 }, { wch: 90 }];
  XLSX.utils.book_append_sheet(wb, detailsWs, "CST Match Details");

  return XLSX.write(wb, { bookType: "xlsx", type: "array" });
}

function collectCstModelsForReport(matches) {
  const values = new Set();
  matches.forEach((match) => {
    [match.factoryModel, match.model].forEach((value) => {
      splitNames(String(value || "")).forEach((item) => values.add(item));
    });
  });
  return [...values];
}

function isModelRepresented(model, reportModels) {
  const normalized = model.toLowerCase();
  return reportModels.some((reportModel) => {
    const candidate = reportModel.toLowerCase();
    return candidate.includes(normalized) || normalized.includes(candidate);
  });
}

async function runAnalysis() {
  if (!state.pdfFiles.length) {
    log("Please load at least one CETR PDF file before analysis.", "ERROR");
    return;
  }
  if (!state.cstFile) {
    log("Please load one CST workbook before analysis.", "ERROR");
    return;
  }
  if (!window.XLSX) {
    log("XLSX library is unavailable. Check network access to the CDN.", "ERROR");
    return;
  }

  state.logs = [];
  resetOutputs();
  els.analyzeButton.disabled = true;
  els.asfState.textContent = "Running";
  els.logState.textContent = "Running";
  log("Analysis started.");

  try {
    setProgress(8, "Reading CST workbook");
    const workbook = await readWorkbook(state.cstFile);
    log(`CST workbook contains ${workbook.SheetNames.length} worksheet(s).`);

    const reports = [];
    for (let index = 0; index < state.pdfFiles.length; index += 1) {
      const file = state.pdfFiles[index];
      setProgress(12 + Math.round((index / state.pdfFiles.length) * 58), `Reading ${file.name}`);
      log(`Extracting text from ${file.name}.`);
      const rawText = await extractPdfText(file);
      const report = analyzeReport(file, rawText);
      log(`Detected ${report.models.length} model candidate(s) in ${file.name}.`);

      const cstMatches = [];
      const modelsNotFoundInCst = [];
      report.models.forEach((model) => {
        const matches = findModelInCst(workbook, model);
        if (!matches.length) {
          modelsNotFoundInCst.push(model);
        }
        matches.forEach((match) => cstMatches.push({ ...match, sourceModel: model }));
      });
      report.cstMatches = cstMatches;
      report.modelsNotFoundInCst = modelsNotFoundInCst;
      report.modelsMissingInCetr = collectCstModelsForReport(cstMatches).filter((model) => !isModelRepresented(model, report.models));
      reports.push(report);
      log(`Matched ${cstMatches.length} CST row(s) for ${file.name}.`);
    }

    setProgress(82, "Building XLSX output");
    const workbookArray = buildSummaryWorkbook(reports, state.cstFile.name);
    state.summaryBlob = new Blob([workbookArray], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    state.logBlob = new Blob([state.logs.join("\n")], { type: "text/plain;charset=utf-8" });

    els.downloadSummaryButton.disabled = false;
    els.downloadLogButton.disabled = false;
    els.asfState.textContent = "Generated";
    els.logState.textContent = "Generated";
    els.asfInfo.textContent = `Analysis Summary generated for ${reports.length} CETR file(s).`;
    els.logInfo.textContent = `${state.logs.length} technical log line(s) generated.`;
    setProgress(100, "Completed");
    log("Analysis completed successfully.");
  } catch (error) {
    state.logBlob = new Blob([state.logs.concat(`${new Date().toISOString()} [ERROR] ${error.stack || error.message}`).join("\n")], { type: "text/plain;charset=utf-8" });
    els.downloadLogButton.disabled = false;
    els.asfState.textContent = "Failed";
    els.logState.textContent = "Generated";
    els.asfInfo.textContent = "Analysis failed. Review status messages and download the log.";
    setProgress(0, "Failed");
    log(error.message || String(error), "ERROR");
  } finally {
    els.analyzeButton.disabled = false;
  }
}

function downloadBlob(blob, fileName) {
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = fileName;
  document.body.append(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}

els.analyzeButton.addEventListener("click", runAnalysis);
els.downloadSummaryButton.addEventListener("click", () => {
  if (state.summaryBlob) downloadBlob(state.summaryBlob, "CE_Test_Report_Analysis_Summary.xlsx");
});
els.downloadLogButton.addEventListener("click", () => {
  const blob = state.logBlob || new Blob([state.logs.join("\n")], { type: "text/plain;charset=utf-8" });
  downloadBlob(blob, "CE_Test_Report_Analyzer.log");
});

renderFileState();
log(`${APP_NAME} v${APP_VERSION} loaded.`);
