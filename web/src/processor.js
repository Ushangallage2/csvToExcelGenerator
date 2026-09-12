const REQUIRED_HEADERS = [
  "Handle",
  "Title",
  "Product Category",
  "Option1 Name",
  "Option1 Value",
  "Option2 Name",
  "Option2 Value",
  "Variant SKU",
];

const VALID_OPTION_TYPES = new Set(["color", "colour", "size", "category", "group", "title"]);

export function cell(row, name) {
  if (!row || name == null) return "";
  const key = Object.keys(row).find((k) => k && k.trim().toLowerCase() === String(name).trim().toLowerCase());
  if (!key) return "";
  const value = row[key];
  return value == null ? "" : String(value).trim();
}

function productError(errorLog, row, metaTitle) {
  return {
    errorLog: errorLog || "",
    handle: cell(row, "Handle"),
    title: metaTitle != null ? String(metaTitle).trim() : cell(row, "Title"),
    productCategory: cell(row, "Product Category"),
    option1Name: cell(row, "Option1 Name"),
    option1Value: cell(row, "Option1 Value"),
    option2Name: cell(row, "Option2 Name"),
    option2Value: cell(row, "Option2 Value"),
    variantSKU: cell(row, "Variant SKU"),
    metaStatus: "",
  };
}

export function validateHeaders(rows) {
  if (!rows.length) return REQUIRED_HEADERS.slice();
  const present = new Set(Object.keys(rows[0]).map((h) => (h || "").trim().toLowerCase()));
  return REQUIRED_HEADERS.filter((h) => !present.has(h.toLowerCase()));
}

export function processShopifyCsv(rows) {
  const missing = validateHeaders(rows);
  if (missing.length) {
    return {
      ok: false,
      message: `Warning: The following required headers are missing from your CSV file: ${missing.join(", ")}. Please update your CSV file headers to include: ${REQUIRED_HEADERS.join(", ")}`,
    };
  }

  const errors = {
    "Invalid - Duplicate SKUs": [],
    "Invalid Options": [],
    "Other Errors": [],
  };
  const successfulRecords = [];
  const existingMetaProductHandles = [];
  const skuSet = new Set();
  let imageEntries = 0;

  const recordsToProcess = [];
  for (const record of rows) {
    const isImageEntry =
      cell(record, "Option1 Name") === "" &&
      cell(record, "Option1 Value") === "" &&
      cell(record, "Option2 Name") === "" &&
      cell(record, "Option2 Value") === "" &&
      cell(record, "Variant SKU") === "";
    if (isImageEntry) imageEntries += 1;
    else recordsToProcess.push(record);
  }

  const handleToRecordsMap = new Map();
  for (const record of recordsToProcess) {
    const handle = cell(record, "Handle");
    if (!handleToRecordsMap.has(handle)) handleToRecordsMap.set(handle, []);
    handleToRecordsMap.get(handle).push(record);
  }

  for (const [handle, records] of handleToRecordsMap.entries()) {
    try {
      const titleDefaultMetaProducts = records.filter(
        (r) =>
          cell(r, "Option1 Name").toLowerCase() === "title" &&
          cell(r, "Option1 Value").toLowerCase() === "default title"
      );

      if (titleDefaultMetaProducts.length > 1) {
        for (const metaRecord of titleDefaultMetaProducts) {
          errors["Other Errors"].push(
            productError(
              "Only one meta product with Option1 Name 'Title' and Option1 Value 'Default Title' is allowed per handle.",
              metaRecord
            )
          );
        }
        continue;
      }

      const metaRecords = records.filter((r) => cell(r, "Title") !== "");

      if (metaRecords.length > 1) {
        for (const metaRecord of metaRecords) {
          errors["Other Errors"].push(
            productError(`Valid title option must have only one record: ${metaRecords.length} found.`, metaRecord)
          );
        }
        continue;
      }

      if (titleDefaultMetaProducts.length || metaRecords.length) {
        if (existingMetaProductHandles.includes(handle)) {
          for (const metaRecord of titleDefaultMetaProducts.length ? titleDefaultMetaProducts : metaRecords) {
            errors["Other Errors"].push(productError(`Meta product handle '${handle}' is not unique.`, metaRecord));
          }
          continue;
        }
        existingMetaProductHandles.push(handle);
      }

      const metaRecord = metaRecords.length ? metaRecords[0] : null;
      let hasMetaProductErrors = false;
      let hasOptionErrors = false;
      const hasNoMetaProduct = metaRecord == null;

      if (
        metaRecord &&
        Object.values(errors)
          .flat()
          .some((error) => error.handle === handle && error.errorLog.includes("Meta product must have a title"))
      ) {
        hasMetaProductErrors = true;
      }

      for (const record of records) {
        const title = cell(record, "Title");
        const option1Name = cell(record, "Option1 Name");
        const option1Value = cell(record, "Option1 Value");
        const option2Name = cell(record, "Option2 Name");
        const option2Value = cell(record, "Option2 Value");
        const sku = cell(record, "Variant SKU");
        if (
          title === "" &&
          option1Name !== "" &&
          option1Value !== "" &&
          option2Name !== "" &&
          option2Value !== "" &&
          sku !== ""
        ) {
          errors["Other Errors"].push(
            productError("This record is suspected as a meta product with missing 'Title' value.", record)
          );
        }
      }

      for (const record of records) {
        const title = cell(record, "Title");
        const sku = cell(record, "Variant SKU");
        const option1Name = cell(record, "Option1 Name");
        const option1Value = cell(record, "Option1 Value");
        const option2Name = cell(record, "Option2 Name");
        const option2Value = cell(record, "Option2 Value");
        const currentRecordErrors = [];

        if (record !== metaRecord) {
          if (option1Name !== "" || option2Name !== "") {
            errors["Invalid Options"].push(productError("Variants cannot define their own option names.", record));
          }
        }

        if (sku === "") {
          const err = productError("Missing SKU", record, metaRecord ? cell(metaRecord, "Title") : "");
          currentRecordErrors.push(err);
          errors["Invalid - Duplicate SKUs"].push(err);
        } else if (skuSet.has(sku)) {
          const err = productError("Duplicate SKU found", record, metaRecord ? cell(metaRecord, "Title") : "");
          currentRecordErrors.push(err);
          errors["Invalid - Duplicate SKUs"].push(err);
        } else {
          skuSet.add(sku);
        }

        if (record === metaRecord) {
          if (title === "") {
            const err = productError("Meta product must have a title", record);
            currentRecordErrors.push(err);
            errors["Other Errors"].push(err);
          }
          if (option1Name === "" && option2Name === "") {
            const err = productError("Meta product cannot have both Option1 Name and Option2 Name empty", record);
            currentRecordErrors.push(err);
            errors["Invalid Options"].push(err);
            hasOptionErrors = true;
          }
          if (option1Name !== "" && !VALID_OPTION_TYPES.has(option1Name.toLowerCase())) {
            const err = productError("Invalid Option1 Name: " + option1Name, record);
            currentRecordErrors.push(err);
            errors["Invalid Options"].push(err);
            hasOptionErrors = true;
          }
          if (option2Name.toLowerCase() === "title") {
            const err = productError("Option2 Name cannot be 'title'", record);
            currentRecordErrors.push(err);
            errors["Invalid Options"].push(err);
            hasOptionErrors = true;
          }
          if (option2Name !== "" && !VALID_OPTION_TYPES.has(option2Name.toLowerCase())) {
            const err = productError("Invalid Option2 Name: " + option2Name, record);
            currentRecordErrors.push(err);
            errors["Invalid Options"].push(err);
            hasOptionErrors = true;
          }
          if (option1Name !== "" && option1Name.toLowerCase() === option2Name.toLowerCase()) {
            const err = productError("Option1 Name and Option2 Name cannot be the same", record);
            currentRecordErrors.push(err);
            errors["Invalid Options"].push(err);
            hasOptionErrors = true;
          }
          if (
            (option1Name.toLowerCase() === "color" && option2Name.toLowerCase() === "colour") ||
            (option1Name.toLowerCase() === "colour" && option2Name.toLowerCase() === "color")
          ) {
            const err = productError(
              "Option names cannot be 'color' and 'colour' simultaneously.  They should be identical.",
              record
            );
            currentRecordErrors.push(err);
            errors["Invalid Options"].push(err);
            hasOptionErrors = true;
          }
          if (option1Name !== "" && option1Value === "") {
            const err = productError("Option1 Value cannot be empty when Option1 Name is present", record);
            currentRecordErrors.push(err);
            errors["Invalid Options"].push(err);
            hasOptionErrors = true;
          }
          if (option2Name !== "" && option2Value === "") {
            const err = productError("Option2 Value cannot be empty when Option2 Name is present", record);
            currentRecordErrors.push(err);
            errors["Invalid Options"].push(err);
            hasOptionErrors = true;
          }
        } else {
          if (title !== "") {
            const err = productError("Variants cannot have a title", record);
            currentRecordErrors.push(err);
            errors["Other Errors"].push(err);
          }
          if (metaRecord) {
            const metaOption1Name = cell(metaRecord, "Option1 Name");
            const metaOption2Name = cell(metaRecord, "Option2 Name");
            if (metaOption1Name !== "" && option1Value === "") {
              const err = productError("Missing value for inherited option: " + metaOption1Name, record);
              currentRecordErrors.push(err);
              errors["Invalid Options"].push(err);
              hasOptionErrors = true;
            }
            if (metaOption2Name !== "" && option2Value === "") {
              const err = productError("Missing value for inherited option: " + metaOption2Name, record);
              currentRecordErrors.push(err);
              errors["Invalid Options"].push(err);
              hasOptionErrors = true;
            }
          }
        }

        if (option1Name.toLowerCase() === "title") {
          if (option1Value.toLowerCase() !== "default title") {
            const err = productError("Option1 Value must be 'Default Title' when Option1 Name is 'title'", record);
            currentRecordErrors.push(err);
            errors["Invalid Options"].push(err);
            hasOptionErrors = true;
          }
          if (option2Name !== "") {
            const err = productError("Option2 Name must be empty when Option1 Name is 'title'", record);
            currentRecordErrors.push(err);
            errors["Invalid Options"].push(err);
            hasOptionErrors = true;
          }
          if (records.length > 1) {
            const err = productError("Variants are not allowed when Option1 Name is 'title'", record);
            currentRecordErrors.push(err);
            errors["Invalid Options"].push(err);
            hasOptionErrors = true;
          }
        }

        let metaStatus = "";
        if (hasNoMetaProduct) metaStatus = "Meta product is missing";
        else if (hasMetaProductErrors || hasOptionErrors) metaStatus = "Meta product has errors";

        if (currentRecordErrors.length) {
          for (const error of currentRecordErrors) error.metaStatus = metaStatus;
        } else {
          successfulRecords.push({ record, metaStatus });
        }
      }
    } catch (rowEx) {
      const msg = `Unexpected error while processing handle '${handle}': ${rowEx.message || rowEx}`;
      if (records.length) errors["Other Errors"].push(productError(msg, records[0]));
    }
  }

  return { ok: true, errors, successfulRecords, imageEntries };
}

export function buildSql(successfulRecords) {
  const skus = [];
  for (const item of successfulRecords) {
    if (item.metaStatus === "Meta product is missing" || item.metaStatus === "Meta product has errors") continue;
    const sku = cell(item.record, "Variant SKU");
    if (sku) skus.push(`('${sku.replace(/'/g, "''")}')`);
  }
  if (!skus.length) return null;
  return (
    "SELECT \n    p.productcode\nFROM \n    (VALUES " +
    skus.join(", ") +
    ") AS p(productcode)\nLEFT JOIN \n    productitem pi ON p.productcode = pi.productcode\nWHERE \n    pi.productcode IS NULL;"
  );
}

export function buildCorrectedRows(successfulRecords) {
  const rows = [];
  for (const item of successfulRecords) {
    if (item.metaStatus === "Meta product is missing" || item.metaStatus === "Meta product has errors") continue;
    const r = item.record;
    rows.push({
      Handle: cell(r, "Handle"),
      Title: cell(r, "Title"),
      "Option1 Name": cell(r, "Option1 Name"),
      "Option1 Value": cell(r, "Option1 Value"),
      "Option2 Name": cell(r, "Option2 Name"),
      "Option2 Value": cell(r, "Option2 Value"),
      "Variant SKU": cell(r, "Variant SKU"),
    });
  }
  return rows;
}

export function hasValidationIssues(result) {
  if (!result.ok) return true;
  const errorCount = Object.values(result.errors).reduce((n, list) => n + list.length, 0);
  if (errorCount > 0) return true;
  return result.successfulRecords.some((r) => r.metaStatus);
}
