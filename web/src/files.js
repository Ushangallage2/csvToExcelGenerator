import Papa from "papaparse";
import JSZip from "jszip";

export function parseCsvText(text) {
  const cleaned = text.replace(/^\uFEFF/, "");
  const firstNl = cleaned.search(/\r\n|\n|\r/);
  let body = cleaned;
  if (firstNl > 0) {
    const header = cleaned.slice(0, firstNl).replace(/,\s*$/, "");
    body = header + cleaned.slice(firstNl);
  }
  const parsed = Papa.parse(body, {
    header: true,
    skipEmptyLines: true,
    transformHeader: (h) => (h || "").trim(),
  });
  if (parsed.errors.length && !parsed.data.length) {
    throw new Error(parsed.errors[0].message || "Could not parse CSV");
  }
  return parsed.data.filter((row) => Object.values(row).some((v) => String(v || "").trim() !== ""));
}

export async function fileToCsvText(file) {
  const name = (file.name || "").toLowerCase();
  if (name.endsWith(".csv") || file.type.includes("csv") || file.type.startsWith("text/")) {
    return file.text();
  }
  if (name.endsWith(".numbers")) {
    return numbersToCsvText(file);
  }
  throw new Error("Please choose a .csv or .numbers file.");
}

export async function numbersToCsvText(file) {
  let zip;
  try {
    zip = await JSZip.loadAsync(await file.arrayBuffer());
  } catch {
    throw new Error(
      "This .numbers file could not be opened in the browser. Export CSV from Apple Numbers, or use the desktop app (which can drive Numbers.app)."
    );
  }

  const csvNames = Object.keys(zip.files).filter((n) => n.toLowerCase().endsWith(".csv") && !zip.files[n].dir);
  if (csvNames.length) {
    return zip.file(csvNames[0]).async("string");
  }

  const tsvNames = Object.keys(zip.files).filter((n) => n.toLowerCase().endsWith(".tsv") && !zip.files[n].dir);
  if (tsvNames.length) {
    const tsv = await zip.file(tsvNames[0]).async("string");
    return tsv.replace(/\t/g, ",");
  }

  throw new Error(
    "Modern Apple Numbers files cannot be fully converted in a browser. Export as CSV from Numbers, or use the macOS desktop app."
  );
}
