
package com.example;

import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVParser;
import org.apache.commons.csv.CSVRecord;
import org.apache.commons.io.input.BOMInputStream;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.*;
import java.util.stream.Collectors;

public class CsvProcessor {

    private static final String[] REQUIRED_HEADERS = { "Handle", "Title", "Product Category", "Option1 Name", "Option1 Value", "Option2 Name", "Option2 Value", "Variant SKU" };

    private static final Set<String> VALID_OPTION_TYPES = new HashSet<>(Arrays.asList("color", "colour", "size", "category", "group", "title"));

    public static void main(String[] args) {
        String inputFilePath = "input.csv"; // Input CSV file
        String outputFilePath = "output.xlsx"; // Output Excel file
        try {
            if (!isFileWritable(outputFilePath)) {
                System.err.println("Error: The output file '" + outputFilePath + "' is open or locked by another process. Please close it and try again.");
                return;
            }

            Map<String, List<CSVRecord>> handleToRecordsMap = new HashMap<>();
            Map<String, List<ProductError>> errors = new HashMap<>();
            errors.put("Invalid - Duplicate SKUs", new ArrayList<>());
            errors.put("Invalid Options", new ArrayList<>());
            errors.put("Other Errors", new ArrayList<>());

            Set<String> skuSet = new HashSet<>();
            List<SuccessfulRecord> successfulRecords = new ArrayList<>(); // Changed to store additional info
            List<CSVRecord> imageEntries = new ArrayList<>();

            try (BOMInputStream bomInputStream = new BOMInputStream(new FileInputStream(inputFilePath));
                 CSVParser parser = new CSVParser(new InputStreamReader(bomInputStream, StandardCharsets.UTF_8),
                         CSVFormat.DEFAULT.withHeader())) {

                // Validate required headers
                Set<String> headersInFile = new HashSet<>(parser.getHeaderMap().keySet());
                List<String> missingHeaders = new ArrayList<>();

                for (String requiredHeader : REQUIRED_HEADERS) {
                    if (!headersInFile.contains(requiredHeader)) {
                        missingHeaders.add(requiredHeader); } }

                if (!missingHeaders.isEmpty()) {
                    System.err.println("Warning: The following required headers are missing from your CSV file: " + missingHeaders);
                    System.err.println("Please update your CSV file headers to include: " + Arrays.toString(REQUIRED_HEADERS));
                    return; //exit
                }

                // Group records by handle and skip image entries
                List<CSVRecord> recordsToProcess = new ArrayList<>();
                for (CSVRecord record : parser) {
                    boolean isImageEntry = record.get("Option1 Name").isEmpty() &&
                            record.get("Option1 Value").isEmpty() &&
                            record.get("Option2 Name").isEmpty() &&
                            record.get("Option2 Value").isEmpty() &&
                            record.get("Variant SKU").isEmpty();

                    if (!isImageEntry) {
                        recordsToProcess.add(record);
                    } else {
                        imageEntries.add(record);
                    }
                }

                for (CSVRecord record : recordsToProcess) {
                    String handle = record.get("Handle");
                    handleToRecordsMap.computeIfAbsent(handle, k -> new ArrayList<>()).add(record);
                }

                // Process each handle group
                for (Map.Entry<String, List<CSVRecord>> entry : handleToRecordsMap.entrySet()) {
                    String handle = entry.getKey();
                    List<CSVRecord> records = entry.getValue();

                    List<CSVRecord> metaRecords = records.stream()
                            .filter(r -> r.get("Title") != null && !r.get("Title").isEmpty())
                            .collect(Collectors.toList());

                    if (metaRecords.size() > 1) {
                        for (CSVRecord metaRecord : metaRecords) {
                            errors.get("Other Errors").add(new ProductError(
                                    "Valid title option must have only one record: " + metaRecords.size() + " found.", metaRecord));
                        }
                        continue;
                    }

                    CSVRecord metaRecord = metaRecords.isEmpty() ? null : metaRecords.get(0);
                    boolean hasMetaProductErrors = false;
                    boolean hasOptionErrors = false; // Track option errors

                    // Check if the handle has no meta product
                    boolean hasNoMetaProduct = metaRecord == null;

                    if (metaRecord != null && errors.values().stream()
                            .flatMap(List::stream)
                            .anyMatch(error -> error.handle.equals(handle) && error.errorLog.contains("Meta product must have a title"))) {
                        hasMetaProductErrors = true;
                    }

                    // Check for suspected meta products (missing title)
                    for (CSVRecord record : records) {
                        String title = getCellValue(record, "Title");
                        String option1Name = getCellValue(record, "Option1 Name");
                        String option1Value = getCellValue(record, "Option1 Value");
                        String option2Name = getCellValue(record, "Option2 Name");
                        String option2Value = getCellValue(record, "Option2 Value");
                        String sku = getCellValue(record, "Variant SKU");
                        if (title.isEmpty() && !option1Name.isEmpty() && !option1Value.isEmpty() &&
                                !option2Name.isEmpty() && !option2Value.isEmpty() && !sku.isEmpty()) {
                            errors.get("Other Errors").add(new ProductError(
                                    "This record is suspected as a meta product with missing 'Title' value.", record));
                        }
                    }

                    // Process each record under this handle
                    for (CSVRecord record : records) {
                        String title = getCellValue(record, "Title");
                        String productCategory = getCellValue(record, "Product Category");
                        String sku = getCellValue(record, "Variant SKU");
                        String option1Name = getCellValue(record, "Option1 Name");
                        String option1Value = getCellValue(record, "Option1 Value");
                        String option2Name = getCellValue(record, "Option2 Name");
                        String option2Value = getCellValue(record, "Option2 Value");

                        // Validate option names for variants
                        if (!record.equals(metaRecord)) {
                            if (!option1Name.isEmpty() || !option2Name.isEmpty()) {
                                errors.get("Invalid Options").add(new ProductError(
                                        "Variants cannot define their own option names.", record));
                            }
                        }

                        // Validate SKU
                        // Inside the inner loop:  for (CSVRecord record : records) {

                        // Collect all errors for this record first
                        List<ProductError> currentRecordErrors = new ArrayList<>();

                        // Validate SKU
                        if (sku.isEmpty()) {
                            currentRecordErrors.add(new ProductError("Missing SKU", record, metaRecord != null ? metaRecord.get("Title") : ""));
                            errors.get("Invalid - Duplicate SKUs").add(currentRecordErrors.get(currentRecordErrors.size() - 1));
                        } else if (skuSet.contains(sku)) {
                            currentRecordErrors.add(new ProductError("Duplicate SKU found", record, metaRecord != null ? metaRecord.get("Title") : ""));
                            errors.get("Invalid - Duplicate SKUs").add(currentRecordErrors.get(currentRecordErrors.size() - 1));
                        } else {
                            skuSet.add(sku);
                        }

                        if (record.equals(metaRecord)) { //Meta Product Validations
                            if (title.isEmpty()) {
                                currentRecordErrors.add(new ProductError("Meta product must have a title", record));
                                errors.get("Other Errors").add(currentRecordErrors.get(currentRecordErrors.size() - 1));
                            }

                            // Meta product cannot have both Option1 Name and Option2 Name empty.
                            if (option1Name.isEmpty() && option2Name.isEmpty()) {
                                currentRecordErrors.add(new ProductError("Meta product cannot have both Option1 Name and Option2 Name empty", record));
                                errors.get("Invalid Options").add(currentRecordErrors.get(currentRecordErrors.size() - 1));
                                hasOptionErrors = true;
                            }

                            if (!option1Name.isEmpty() && !VALID_OPTION_TYPES.contains(option1Name.toLowerCase())) {
                                currentRecordErrors.add(new ProductError("Invalid Option1 Name: " + option1Name, record));
                                errors.get("Invalid Options").add(currentRecordErrors.get(currentRecordErrors.size() - 1));
                                hasOptionErrors = true;
                            }

                            if (option2Name.equalsIgnoreCase("title")) {
                                currentRecordErrors.add(new ProductError("Option2 Name cannot be 'title'", record));
                                errors.get("Invalid Options").add(currentRecordErrors.get(currentRecordErrors.size() - 1));
                                hasOptionErrors = true;
                            }

                            if (!option2Name.isEmpty() && !VALID_OPTION_TYPES.contains(option2Name.toLowerCase())) {
                                currentRecordErrors.add(new ProductError("Invalid Option2 Name: " + option2Name, record));
                                errors.get("Invalid Options").add(currentRecordErrors.get(currentRecordErrors.size() - 1));
                                hasOptionErrors = true;
                            }

                            if (!option1Name.isEmpty() && option1Name.equalsIgnoreCase(option2Name)) {
                                currentRecordErrors.add(new ProductError("Option1 Name and Option2 Name cannot be the same", record));
                                errors.get("Invalid Options").add(currentRecordErrors.get(currentRecordErrors.size() - 1));
                                hasOptionErrors = true;
                            }

                            // New Validation: Color vs. Colour
                            if (option1Name.equalsIgnoreCase("color") && option2Name.equalsIgnoreCase("colour") ||
                                    option1Name.equalsIgnoreCase("colour") && option2Name.equalsIgnoreCase("color")) {
                                currentRecordErrors.add(new ProductError("Option names cannot be 'color' and 'colour' simultaneously.  They should be identical.", record));
                                errors.get("Invalid Options").add(currentRecordErrors.get(currentRecordErrors.size() - 1));
                                hasOptionErrors = true;
                            }
                            if (!option1Name.isEmpty() && option1Value.isEmpty()) {
                                currentRecordErrors.add(new ProductError("Option1 Value cannot be empty when Option1 Name is present", record));
                                errors.get("Invalid Options").add(currentRecordErrors.get(currentRecordErrors.size() - 1));
                                hasOptionErrors = true;
                            }
                            if (!option2Name.isEmpty() && option2Value.isEmpty()) {
                                currentRecordErrors.add(new ProductError("Option2 Value cannot be empty when Option2 Name is present", record));
                                errors.get("Invalid Options").add(currentRecordErrors.get(currentRecordErrors.size() - 1));
                                hasOptionErrors = true;
                            }
                        } else { //Variant Product Validations
                            if (!title.isEmpty()) {
                                currentRecordErrors.add(new ProductError("Variants cannot have a title", record));
                                errors.get("Other Errors").add(currentRecordErrors.get(currentRecordErrors.size() - 1));
                            }

                            // Check if the variant has values for meta product options
                            if (metaRecord != null) {
                                String metaOption1Name = metaRecord.get("Option1 Name");
                                String metaOption2Name = metaRecord.get("Option2 Name");

                                if (!metaOption1Name.isEmpty() && option1Value.isEmpty()) {
                                    currentRecordErrors.add(new ProductError("Missing value for inherited option: " + metaOption1Name, record));
                                    errors.get("Invalid Options").add(currentRecordErrors.get(currentRecordErrors.size() - 1));
                                    hasOptionErrors = true;
                                }
                                if (!metaOption2Name.isEmpty() && option2Value.isEmpty()) {
                                    currentRecordErrors.add(new ProductError("Missing value for inherited option: " + metaOption2Name, record));
                                    errors.get("Invalid Options").add(currentRecordErrors.get(currentRecordErrors.size() - 1));
                                    hasOptionErrors = true;
                                }
                            }
                        }

                        // Custom Validations
                        if (option1Name.equalsIgnoreCase("title")) {
                            if (!option1Value.equalsIgnoreCase("Default Title")) {
                                currentRecordErrors.add(new ProductError("Option1 Value must be 'Default Title' when Option1 Name is 'title'", record));
                                errors.get("Invalid Options").add(currentRecordErrors.get(currentRecordErrors.size() - 1));
                                hasOptionErrors = true;
                            }
                            if (!option2Name.isEmpty()) {
                                currentRecordErrors.add(new ProductError("Option2 Name must be empty when Option1 Name is 'title'", record));
                                errors.get("Invalid Options").add(currentRecordErrors.get(currentRecordErrors.size() - 1));
                                hasOptionErrors = true;
                            }

                            // Check for variants (more than one record for the handle)
                            if (records.size() > 1) {
                                currentRecordErrors.add(new ProductError("Variants are not allowed when Option1 Name is 'title'", record));
                                errors.get("Invalid Options").add(currentRecordErrors.get(currentRecordErrors.size() - 1));
                                hasOptionErrors = true;
                            }
                        }

                        String metaStatus = "";
                        if (hasNoMetaProduct) {
                            metaStatus = "Meta product is missing";
                        } else if (hasMetaProductErrors || hasOptionErrors) {
                            metaStatus = "Meta product has errors";
                        }

                        //  Determine if the variant itself has ANY errors (including those found above)
                        boolean hasVariantErrors = !currentRecordErrors.isEmpty();

                        if (hasVariantErrors) {
                            //Add meta status to all generated errors
                            for (ProductError error : currentRecordErrors) {
                                error.metaStatus = metaStatus;
                            }
                        } else {
                            // If the variant has no errors of its own, but the meta product is missing or has errors,
                            // then add it to successful records with the meta status.
                            if (!metaStatus.isEmpty()) {
                                successfulRecords.add(new SuccessfulRecord(record, metaStatus));
                            } else {
                                // Otherwise, it's a completely successful record.
                                successfulRecords.add(new SuccessfulRecord(record, ""));
                            }
                        }
                    }

                }
            }

            System.out.println("Skipped image entries: " + imageEntries.size());

            writeErrorsToExcel(outputFilePath, errors);
            writeSuccessfulRecordsToExcel(outputFilePath, successfulRecords);
            System.out.println("Processing completed. Errors written to: " + outputFilePath);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static void writeErrorsToExcel(String outputFilePath, Map<String, List<ProductError>> errors) throws IOException {
        try (Workbook workbook = new XSSFWorkbook()) {
            for (Map.Entry<String, List<ProductError>> entry : errors.entrySet()) {
                writeErrorsToSheet(workbook, entry.getKey(), entry.getValue());
            }
            try (FileOutputStream outputStream = new FileOutputStream(outputFilePath)) {
                workbook.write(outputStream);
            }
        }
    }

    private static void writeErrorsToSheet(Workbook workbook, String sheetName, List<ProductError> productErrors) {
        Sheet sheet = workbook.createSheet(sheetName);
        Row countRow = sheet.createRow(0);
        countRow.createCell(0).setCellValue("Count of " + sheetName + ": " + productErrors.size());

        Row headerRow = sheet.createRow(1);
        String[] headers = {"Error Log", "Handle", "Title", "Product Category", "Option 1 Name", "Option 1 Value", "Option 2 Name", "Option 2 Value", "Variant SKU", "Meta Status"};
        for (int i = 0; i < headers.length; i++) {
            headerRow.createCell(i).setCellValue(headers[i]);
        }

        int rowNum = 2;
        for (ProductError error : productErrors) {
            Row row = sheet.createRow(rowNum++);
            row.createCell(0).setCellValue(error.errorLog);
            row.createCell(1).setCellValue(error.handle);
            row.createCell(2).setCellValue(error.title);
            row.createCell(3).setCellValue(error.productCategory);
            row.createCell(4).setCellValue(error.option1Name);
            row.createCell(5).setCellValue(error.option1Value);
            row.createCell(6).setCellValue(error.option2Name);
            row.createCell(7).setCellValue(error.option2Value);
            row.createCell(8).setCellValue(error.variantSKU);
            row.createCell(9).setCellValue(error.metaStatus != null ? error.metaStatus : "");
        }
    }

    private static void writeSuccessfulRecordsToExcel(String outputFilePath, List<SuccessfulRecord> successfulRecords) throws IOException {
        try (FileInputStream fileInputStream = new FileInputStream(outputFilePath);
             Workbook workbook = new XSSFWorkbook(fileInputStream)) {
            Sheet successSheet = workbook.createSheet("Success");
            Row countRow = successSheet.createRow(0);
            countRow.createCell(0).setCellValue("Count of Successful Records: " + successfulRecords.size());

            Row headerRow = successSheet.createRow(1);
            String[] headers = {"Handle", "Title", "Product Category", "Option1 Name", "Option1 Value", "Option2 Name", "Option2 Value", "Variant SKU", "Meta Status"};
            for (int i = 0; i < headers.length; i++) {
                headerRow.createCell(i).setCellValue(headers[i]);
            }

            int rowNum = 2;
            for (SuccessfulRecord successfulRecord : successfulRecords) {
                CSVRecord record = successfulRecord.record;
                Row row = successSheet.createRow(rowNum++);
                row.createCell(0).setCellValue(record.get("Handle"));
                row.createCell(1).setCellValue(record.get("Title"));
                row.createCell(2).setCellValue(record.get("Product Category"));
                row.createCell(3).setCellValue(record.get("Option1 Name"));
                row.createCell(4).setCellValue(record.get("Option1 Value"));
                row.createCell(5).setCellValue(record.get("Option2 Name"));
                row.createCell(6).setCellValue(record.get("Option2 Value"));
                row.createCell(7).setCellValue(record.get("Variant SKU"));
                row.createCell(8).setCellValue(successfulRecord.metaStatus);
            }

            try (FileOutputStream outputStream = new FileOutputStream(outputFilePath)) {
                workbook.write(outputStream);
            }
        }
    }

    private static boolean isFileWritable(String filePath) {
        File file = new File(filePath);
        return !file.exists() || (Files.isWritable(Paths.get(filePath)) && !isFileLocked(String.valueOf(file)));
    }

    private static boolean isFileLocked(String file) {
        try (RandomAccessFile raf = new RandomAccessFile(file, "rw")) {
            return false;
        } catch (IOException e) {
            return true;
        }
    }

    private static String getCellValue(CSVRecord record, String headerName) {
        try {
            return record.get(headerName);
        } catch (IllegalArgumentException e) {
            return ""; // Handle missing header gracefully
        }
    }

    static class ProductError {
        String errorLog;
        String handle;
        String title;
        String productCategory;
        String option1Name;
        String option1Value;
        String option2Name;
        String option2Value;
        String variantSKU;
        String metaStatus;

        public ProductError(String errorLog, CSVRecord record) {
            this.errorLog = errorLog;
            this.handle = record.get("Handle");
            this.title = record.get("Title");
            this.productCategory = record.get("Product Category");
            this.option1Name = record.get("Option1 Name");
            this.option1Value = record.get("Option1 Value");
            this.option2Name = record.get("Option2 Name");
            this.option2Value = record.get("Option2 Value");
            this.variantSKU = record.get("Variant SKU");
            this.metaStatus = null;
        }

        public ProductError(String errorLog, CSVRecord record, String metaTitle) {
            this(errorLog, record);
            this.title = metaTitle;
        }
    }

    static class SuccessfulRecord {
        CSVRecord record;
        String metaStatus;

        public SuccessfulRecord(CSVRecord record, String metaStatus) {
            this.record = record;
            this.metaStatus = metaStatus;
        }
    }
}