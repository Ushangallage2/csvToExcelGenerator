import "./style.css";
import { parseCsvText, fileToCsvText, numbersToCsvText } from "./files.js";
import { processShopifyCsv, buildSql, buildCorrectedRows, hasValidationIssues } from "./processor.js";
import { processVariationUpload, processProductUpload } from "./templates.js";
import { buildValidationWorkbook, buildCorrectedWorkbook, downloadXlsx, downloadText, downloadCsv } from "./excel.js";

const logEl = document.getElementById("log");
const statusA = document.getElementById("status-a");
const statusB = document.getElementById("status-b");
const hint = document.getElementById("hint");
const hintArrow = document.getElementById("hint-arrow");
const guidance = document.getElementById("guidance");

let csvRows = [];
let csvName = "input.csv";
let attempt = 1;
let lastWorkbook = null;
let lastWorkbookName = "";
let lastResult = null;
let lastTemplate = null;

function log(kind, message) {
  const prefix = kind === "error" ? "Error: " : "Info: ";
  logEl.textContent += prefix + message + "\n";
  logEl.scrollTop = logEl.scrollHeight;
}

hint.addEventListener("click", () => {
  const open = guidance.hasAttribute("hidden");
  guidance.toggleAttribute("hidden", !open);
  hintArrow.textContent = open ? "▲" : "▼";
});

async function loadCsvFiles(fileList) {
  const files = [...fileList];
  if (!files.length) return;
  const file = files[0];
  try {
    const text = await fileToCsvText(file);
    csvRows = parseCsvText(text);
    csvName = file.name.replace(/\.(csv|numbers)$/i, "");
    statusA.textContent = `${files.length} file loaded: ${file.name} (${csvRows.length} rows)`;
    log("info", `Loaded ${file.name}`);
  } catch (e) {
    log("error", e.message || String(e));
  }
}

document.getElementById("select-csv").addEventListener("click", () => {
  document.getElementById("csv-input").click();
});
document.getElementById("csv-input").addEventListener("change", (e) => loadCsvFiles(e.target.files));

document.getElementById("process-csv").addEventListener("click", async () => {
  if (!csvRows.length) {
    log("error", "Please select CSV files first.");
    return;
  }
  const result = processShopifyCsv(csvRows);
  if (!result.ok) {
    log("error", result.message);
    lastResult = null;
    lastWorkbook = null;
    return;
  }
  lastResult = result;
  lastWorkbook = await buildValidationWorkbook(result);
  lastWorkbookName = `${csvName}_attempt_${attempt}.xlsx`;
  attempt += 1;
  log("info", `Output ready: ${lastWorkbookName}`);
  if (hasValidationIssues(result)) log("error", `Validation issues found — review the Excel report: ${csvName}`);
  else log("info", `No validation errors: ${csvName}`);
  log("info", "All files processed.");
});

document.getElementById("view-excel").addEventListener("click", async () => {
  if (!lastWorkbook) {
    log("error", "No processed Excel files available. Process a CSV first.");
    return;
  }
  await downloadXlsx(lastWorkbook, lastWorkbookName);
  log("info", `Opened/downloaded: ${lastWorkbookName}`);
});

document.getElementById("save-excel").addEventListener("click", async () => {
  if (!lastWorkbook) {
    log("error", "No Excel files have been processed yet.");
    return;
  }
  await downloadXlsx(lastWorkbook, lastWorkbookName);
  log("info", `Excel file saved: ${lastWorkbookName}`);
});

document.getElementById("corrected").addEventListener("click", async () => {
  if (!lastResult) {
    log("error", "No processed Excel files available. Process a CSV first.");
    return;
  }
  const rows = buildCorrectedRows(lastResult.successfulRecords);
  const wb = await buildCorrectedWorkbook(rows);
  await downloadXlsx(wb, "CorrectedOutput.xlsx");
  log("info", "Corrected output saved successfully: CorrectedOutput.xlsx");
});

document.getElementById("sql").addEventListener("click", () => {
  if (!lastResult) {
    log("error", "No processed Excel files available. Process a CSV first.");
    return;
  }
  const query = buildSql(lastResult.successfulRecords);
  if (!query) {
    log("error", "No valid Variant SKUs found to build SQL.");
    return;
  }
  downloadText(query, "query.sql", "application/sql");
  log("info", "SQL query saved successfully: query.sql");
});

document.getElementById("process-template").addEventListener("click", () => {
  const kind = document.getElementById("template-type").value;
  if (!kind) {
    log("error", "Choose a Variation or Product template first.");
    return;
  }
  if (!csvRows.length) {
    log("error", "No file selected.");
    return;
  }
  if (kind === "variation") {
    lastTemplate = processVariationUpload(csvRows);
  } else {
    lastTemplate = processProductUpload(csvRows);
  }
  log(lastTemplate.ok ? "info" : "error", lastTemplate.message);
  if (lastTemplate.ok) statusB.textContent = `${kind} template processed`;
});

document.getElementById("save-template").addEventListener("click", () => {
  if (!lastTemplate || !lastTemplate.ok) {
    log("error", "Process a template first.");
    return;
  }
  if (lastTemplate.csv) downloadCsv(lastTemplate.csv, lastTemplate.filename);
  if (lastTemplate.files) lastTemplate.files.forEach((f) => downloadCsv(f.csv, f.filename));
  log("info", "Template file(s) downloaded.");
});

document.getElementById("convert-numbers").addEventListener("click", () => {
  document.getElementById("numbers-input").click();
});
document.getElementById("numbers-input").addEventListener("change", async (e) => {
  const file = e.target.files?.[0];
  if (!file) return;
  try {
    const text = await numbersToCsvText(file);
    const name = file.name.replace(/\.numbers$/i, "") + ".csv";
    downloadCsv(text, name);
    csvRows = parseCsvText(text);
    csvName = file.name.replace(/\.numbers$/i, "");
    statusA.textContent = `Converted and loaded: ${name} (${csvRows.length} rows)`;
    log("info", `Conversion successful! Saved to: ${name}`);
  } catch (err) {
    log("error", "Conversion failed: " + (err.message || err));
  }
});

document.getElementById("clear-all").addEventListener("click", () => {
  csvRows = [];
  lastWorkbook = null;
  lastResult = null;
  lastTemplate = null;
  csvName = "input.csv";
  statusA.textContent = "No CSV file selected";
  statusB.textContent = "No Variation upload file selected";
  logEl.textContent = "";
  document.getElementById("csv-input").value = "";
  document.getElementById("numbers-input").value = "";
  log("info", "Cleared.");
});
