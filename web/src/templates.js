function cell(row, name) {
  const key = Object.keys(row).find((k) => k && k.trim().toLowerCase() === name.toLowerCase());
  if (!key) return "";
  const value = row[key];
  return value == null ? "" : String(value).trim();
}

function csvEscape(value) {
  const s = value == null ? "" : String(value);
  if (/[",\n\r]/.test(s)) return `"${s.replace(/"/g, '""')}"`;
  return s;
}

function toCsv(headers, rows) {
  const lines = [headers.join(",")];
  for (const row of rows) {
    lines.push(headers.map((h) => csvEscape(row[h])).join(","));
  }
  return lines.join("\n") + "\n";
}

export function processVariationUpload(rows) {
  const required = ["variation_name", "option1", "option2", "product_code"];
  const headerMap = rows[0] ? Object.keys(rows[0]).map((h) => h.trim().toLowerCase()) : [];
  for (const req of required) {
    if (!headerMap.includes(req)) {
      return { ok: false, message: `Missing required header: ${req}` };
    }
  }

  const messages = [];
  const validRecords = [];
  const variationNames = new Set();
  const productCodes = new Set();
  const duplicateVariationNames = new Set();
  const duplicateProductCodes = new Set();

  let rowNum = 2;
  for (const record of rows) {
    const variationName = cell(record, "variation_name");
    const option1 = cell(record, "option1");
    const option2 = cell(record, "option2");
    const productCode = cell(record, "product_code");

    if (!variationName) {
      rowNum += 1;
      continue;
    }
    if (!option1 && !option2) {
      messages.push(`Row ${rowNum}: Must have at least option1 or option2 for variation_name: ${variationName}`);
      rowNum += 1;
      continue;
    }
    if (option2 && !option1) {
      messages.push(`Row ${rowNum}: Has option2 but missing option1 for variation_name: ${variationName}`);
      rowNum += 1;
      continue;
    }
    if (variationNames.has(variationName)) {
      duplicateVariationNames.add(variationName);
      rowNum += 1;
      continue;
    }
    variationNames.add(variationName);
    if (productCode && productCodes.has(productCode)) {
      duplicateProductCodes.add(productCode);
      rowNum += 1;
      continue;
    }
    if (productCode) productCodes.add(productCode);
    validRecords.push({ variation_name: variationName, option1, option2, meta_product_code: productCode });
    rowNum += 1;
  }

  if (duplicateVariationNames.size) messages.push(`Duplicate variation_name(s): ${[...duplicateVariationNames].join(", ")}`);
  if (duplicateProductCodes.size) messages.push(`Duplicate product_code(s): ${[...duplicateProductCodes].join(", ")}`);

  if (!validRecords.length) {
    messages.push("No valid records to write.");
    return { ok: false, message: messages.join("\n") };
  }

  const csv = toCsv(["variation_name", "option1", "option2", "meta_product_code"], validRecords);
  messages.push("Variation upload processed. Use Save Template to download.");
  return { ok: true, filename: "variation_upload_processed.csv", csv, message: messages.join("\n") };
}

export function processProductUpload(rows) {
  const exportHeaders = ["variation_name", "option1", "option2", "product_code"];
  const headerMap = rows[0] ? Object.keys(rows[0]).map((h) => h.trim().toLowerCase()) : [];
  for (const req of exportHeaders) {
    if (!headerMap.includes(req)) {
      return { ok: false, message: `Missing required header: ${req}` };
    }
  }

  const allGroups = [];
  const groupNames = [];
  let currentGroup = [];
  for (const record of rows) {
    const blank = exportHeaders.every((h) => cell(record, h) === "");
    if (blank) continue;
    const variationName = cell(record, "variation_name");
    if (variationName) {
      if (currentGroup.length) {
        allGroups.push(currentGroup);
        currentGroup = [];
      }
      groupNames.push(variationName);
    }
    currentGroup.push(record);
  }
  if (currentGroup.length) allGroups.push(currentGroup);

  const variationNameToGroups = new Map();
  groupNames.forEach((name, i) => {
    if (!variationNameToGroups.has(name)) variationNameToGroups.set(name, []);
    variationNameToGroups.get(name).push(i);
  });
  const productCodeToGroups = new Map();
  allGroups.forEach((group, i) => {
    for (const record of group) {
      const productCode = cell(record, "product_code");
      if (!productCode) continue;
      if (!productCodeToGroups.has(productCode)) productCodeToGroups.set(productCode, []);
      productCodeToGroups.get(productCode).push(i);
    }
  });

  const invalidGroupIndexes = new Set();
  for (const indexes of variationNameToGroups.values()) {
    if (indexes.length > 1) indexes.forEach((i) => invalidGroupIndexes.add(i));
  }
  for (const indexes of productCodeToGroups.values()) {
    if (indexes.length > 1) indexes.forEach((i) => invalidGroupIndexes.add(i));
  }
  allGroups.forEach((group, i) => {
    const localProductCodes = new Set();
    group.forEach((record, j) => {
      const productCode = cell(record, "product_code");
      const option1 = cell(record, "option1");
      if (j === 0 && option1 === "") invalidGroupIndexes.add(i);
      if (!productCode || localProductCodes.has(productCode)) invalidGroupIndexes.add(i);
      else localProductCodes.add(productCode);
    });
  });

  const validGroups = [];
  const invalidGroups = [];
  allGroups.forEach((group, i) => {
    (invalidGroupIndexes.has(i) ? invalidGroups : validGroups).push(group);
  });

  const messages = [];
  const files = [];
  const toRows = (groups) =>
    groups.flatMap((group) =>
      group.map((rec) => ({
        variation_name: cell(rec, "variation_name"),
        option1: cell(rec, "option1"),
        option2: cell(rec, "option2"),
        product_code: cell(rec, "product_code"),
      }))
    );

  if (validGroups.length) {
    files.push({
      filename: "product_upload_processed.csv",
      csv: toCsv(exportHeaders, toRows(validGroups)),
    });
  } else {
    messages.push("No valid product groups to write.");
  }
  if (invalidGroups.length) {
    files.push({
      filename: "invalid.csv",
      csv: toCsv(exportHeaders, toRows(invalidGroups)),
    });
    messages.push("Invalid records written to invalid.csv");
  }
  if (files.length) messages.push("Product upload processed. Use Save Template to download.");
  return { ok: files.length > 0, files, message: messages.join("\n") };
}
