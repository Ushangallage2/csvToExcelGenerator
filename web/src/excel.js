import ExcelJS from "exceljs";

function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  a.click();
  URL.revokeObjectURL(url);
}

export async function workbookToBlob(workbook) {
  const buffer = await workbook.xlsx.writeBuffer();
  return new Blob([buffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
}

export async function buildValidationWorkbook(result) {
  const workbook = new ExcelJS.Workbook();
  const errorHeaders = [
    "Error Log",
    "Handle",
    "Title",
    "Product Category",
    "Option 1 Name",
    "Option 1 Value",
    "Option 2 Name",
    "Option 2 Value",
    "Variant SKU",
    "Meta Status",
  ];

  for (const [sheetName, productErrors] of Object.entries(result.errors)) {
    const sheet = workbook.addWorksheet(sheetName);
    sheet.addRow([`Count of ${sheetName}: ${productErrors.length}`]);
    sheet.addRow(errorHeaders);
    for (const error of productErrors) {
      sheet.addRow([
        error.errorLog || "",
        error.handle || "",
        error.title || "",
        error.productCategory || "",
        error.option1Name || "",
        error.option1Value || "",
        error.option2Name || "",
        error.option2Value || "",
        error.variantSKU || "",
        error.metaStatus || "",
      ]);
    }
  }

  const success = workbook.addWorksheet("Success");
  success.addRow([`Count of Successful Records: ${result.successfulRecords.length}`]);
  success.addRow([
    "Handle",
    "Title",
    "Product Category",
    "Option1 Name",
    "Option1 Value",
    "Option2 Name",
    "Option2 Value",
    "Variant SKU",
    "Meta Status",
  ]);
  for (const item of result.successfulRecords) {
    const r = item.record;
    const c = (name) => {
      const key = Object.keys(r).find((k) => k && k.trim().toLowerCase() === name.toLowerCase());
      const value = key ? r[key] : "";
      return value == null ? "" : String(value);
    };
    success.addRow([
      c("Handle"),
      c("Title"),
      c("Product Category"),
      c("Option1 Name"),
      c("Option1 Value"),
      c("Option2 Name"),
      c("Option2 Value"),
      c("Variant SKU"),
      item.metaStatus || "",
    ]);
  }
  return workbook;
}

export async function buildCorrectedWorkbook(rows) {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("Filtered Success");
  const headers = [
    "Handle",
    "Title",
    "Option1 Name",
    "Option1 Value",
    "Option2 Name",
    "Option2 Value",
    "Variant SKU",
  ];
  const dest = { Handle: 1, Title: 2, "Option1 Name": 9, "Option1 Value": 10, "Option2 Name": 12, "Option2 Value": 13, "Variant SKU": 18 };
  const headerRow = sheet.getRow(1);
  for (const [name, col] of Object.entries(dest)) {
    headerRow.getCell(col).value = name;
  }
  headerRow.commit();
  rows.forEach((row, i) => {
    const excelRow = sheet.getRow(i + 2);
    for (const name of headers) {
      excelRow.getCell(dest[name]).value = row[name] || "";
    }
    excelRow.commit();
  });
  return workbook;
}

export async function downloadXlsx(workbook, filename) {
  downloadBlob(await workbookToBlob(workbook), filename);
}

export function downloadText(text, filename, mime = "text/plain") {
  downloadBlob(new Blob([text], { type: mime }), filename);
}

export function downloadCsv(text, filename) {
  downloadText(text, filename, "text/csv;charset=utf-8");
}
