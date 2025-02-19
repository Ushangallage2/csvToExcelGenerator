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

                String[] requiredHeaders = {"Handle", "Title", "Product Category", "Option1 Name", "Option1 Value", "Option2 Name", "Option2 Value", "Variant SKU"};

                // Validate required headers
                for (String header : requiredHeaders) {
                    if (!parser.getHeaderMap().containsKey(header)) {
                        errors.get("Other Errors").add(new ProductError("Missing required header: " + header, null));
                    }
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

                    if (metaRecord == null) {
                        // Handle cases where there is no meta product
                        for (CSVRecord record : records) {
                            String sku = getCellValue(record, "Variant SKU");
                            String title = getCellValue(record, "Title");
                            String option1Value = getCellValue(record, "Option1 Value");
                            String option2Value = getCellValue(record, "Option2 Value");

                            // Check if the variant can be added to the "Success" sheet
                            if (!sku.isEmpty() && !handle.isEmpty() && title.isEmpty() &&
                                    (!option1Value.isEmpty() || !option2Value.isEmpty())) {
                                successfulRecords.add(new SuccessfulRecord(record, "Meta product is missing"));
                            } else {
                                errors.get("Other Errors").add(new ProductError("Missing meta product title", record));
                            }
                        }
                        continue;
                    }

                    String metaTitle = metaRecord.get("Title");
                    String metaOption1Name = metaRecord.get("Option1 Name");
                    String metaOption2Name = metaRecord.get("Option2 Name");

                    // Check for valid option names in meta product
                    if (metaOption1Name.isEmpty() && metaOption2Name.isEmpty()) {
                        errors.get("Other Errors").add(new ProductError("Meta product must have at least one option name.", metaRecord));
                    }

                    // Check for duplicate option names
                    if (!metaOption1Name.isEmpty() && metaOption1Name.equalsIgnoreCase(metaOption2Name)) {
                        errors.get("Invalid Options").add(new ProductError("Duplicate option names found in meta product: " + metaOption1Name, metaRecord));
                    }

                    // Check for special 'title' option in meta product
                    boolean hasTitleOption = false;
                    if (metaOption1Name.equalsIgnoreCase("title")) {
                        hasTitleOption = true;

                        // Validate 'title' option rules
                        String option1Value = metaRecord.get("Option1 Value");
                        if (!option1Value.equalsIgnoreCase("default title")) {
                            errors.get("Invalid Options").add(new ProductError("Option1 Value must be 'Default Title' when Option1 Name is 'title'", metaRecord));
                        }

                        if (!metaOption2Name.isEmpty()) {
                            errors.get("Invalid Options").add(new ProductError("Option2 Name must be empty when Option1 Name is 'title'", metaRecord));
                        }

                        if (records.size() > 1) {
                            for (CSVRecord record : records) {
                                if (!record.equals(metaRecord)) {
                                    errors.get("Invalid Options").add(new ProductError("Variants are not allowed when Option1 Name is 'title'", record));
                                }
                            }
                        }
                    } else if (metaOption2Name.equalsIgnoreCase("title")) {
                        errors.get("Invalid Options").add(new ProductError("'title' option must be in Option1 Name", metaRecord));
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

                        // Validate option names
                        if (!option1Name.isEmpty() && !VALID_OPTION_TYPES.contains(option1Name.toLowerCase())) {
                            errors.get("Invalid Options").add(new ProductError("Option 1 Name must be 'Color', 'Size', 'Brand', 'Category', or 'Title' ", record));
                        }
                        if (!option2Name.isEmpty() && !VALID_OPTION_TYPES.contains(option2Name.toLowerCase())) {
                            errors.get("Invalid Options").add(new ProductError("Option 2 Name must be 'Color', 'Size', 'Brand', 'Category', or 'Title'; ", record));
                        }

                        // Validate SKU
                        if (sku.isEmpty()) {
                            errors.get("Invalid - Duplicate SKUs").add(new ProductError("Missing SKU", record, metaTitle));
                        } else if (skuSet.contains(sku)) {
                            errors.get("Invalid - Duplicate SKUs").add(new ProductError("Duplicate SKU found", record, metaTitle));
                        } else {
                            skuSet.add(sku);
                        }

                        if (record.equals(metaRecord)) {
                            if (title.isEmpty()) {
                                errors.get("Other Errors").add(new ProductError("Meta product must have a title", record));
                            }
                        } else {
                            if (!title.isEmpty()) {
                                errors.get("Other Errors").add(new ProductError("Variants cannot have a title", record));
                            }

                            // Check if the variant has values for meta product options
                            if (!metaOption1Name.isEmpty() && option1Value.isEmpty()) {
                                errors.get("Invalid Options").add(new ProductError(
                                        "Missing value for inherited option: " + metaOption1Name, record));
                            }
                            if (!metaOption2Name.isEmpty() && option2Value.isEmpty()) {
                                errors.get("Invalid Options").add(new ProductError(
                                        "Missing value for inherited option: " + metaOption2Name, record));
                            }
                        }

                        // Check if the meta product has errors
                        boolean metaHasErrors = errors.values().stream()
                                .anyMatch(list -> list.stream().anyMatch(e -> e.handle.equals(handle) && e.record == metaRecord));

                        // If no errors were added, consider the record successful
                        if (errors.values().stream().noneMatch(list -> list.stream().anyMatch(e -> e.handle.equals(record.get("Handle"))))) {
                            successfulRecords.add(new SuccessfulRecord(record, metaHasErrors ? "Meta product has errors" : ""));
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
        String[] headers = {"Error Log", "Handle", "Title", "Product Category", "Option 1 Name", "Option 1 Value", "Option 2 Name", "Option 2 Value", "Variant SKU"};
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
        }
    }

    private static void writeSuccessfulRecordsToExcel(String outputFilePath, List<SuccessfulRecord> successfulRecords) throws IOException {
        try (FileInputStream fileInputStream = new FileInputStream(outputFilePath);
             Workbook workbook = new XSSFWorkbook(fileInputStream)) {
            Sheet successSheet = workbook.createSheet("Success");
            Row countRow = successSheet.createRow(0);
            countRow.createCell(0).setCellValue("Count of Successful Records: " + successfulRecords.size());

            Row headerRow = successSheet.createRow(1);
            String[] headers = {"Handle", "Title", "Product Category", "Option1 Name", "Option1 Value", "Option2 Name", "Option2 Value", "Variant SKU", "Meta Product Status"}; // Added new header
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
                row.createCell(8).setCellValue(successfulRecord.metaStatus); // Add the meta product status
            }

            try (FileOutputStream outputStream = new FileOutputStream(outputFilePath)) {
                workbook.write(outputStream);
            }
        }
    }

    private static boolean isFileWritable(String filePath) {
        File file = new File(filePath);
        return !file.exists() || (Files.isWritable(Paths.get(filePath)) && !isFileLocked(file));
    }

    private static boolean isFileLocked(File file) {
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

    // Helper class to store successful records with meta product status
    private static class SuccessfulRecord {
        CSVRecord record;
        String metaStatus; // "Meta product is missing" or "Meta product has errors"

        public SuccessfulRecord(CSVRecord record, String metaStatus) {
            this.record = record;
            this.metaStatus = metaStatus;
        }
    }

    public static class ProductError {
        String errorLog;
        String handle;
        String title;
        String productCategory;
        String option1Name;
        String option1Value;
        String option2Name;
        String option2Value;
        String variantSKU;
        CSVRecord record;

        public ProductError(String errorLog, CSVRecord record) {
            this.errorLog = errorLog;
            this.record = record;
            if (record != null) {
                this.handle = record.get("Handle");
                this.title = record.get("Title");
                this.productCategory = record.get("Product Category");
                this.option1Name = record.get("Option1 Name");
                this.option1Value = record.get("Option1 Value");
                this.option2Name = record.get("Option2 Name");
                this.option2Value = record.get("Option2 Value");
                this.variantSKU = record.get("Variant SKU");
            } else {
                this.handle = "";
                this.title = "";
                this.productCategory = "";
                this.option1Name = "";
                this.option1Value = "";
                this.option2Name = "";
                this.option2Value = "";
                this.variantSKU = "";
            }
        }

        public ProductError(String errorLog, CSVRecord record, String metaTitle) {
            this(errorLog, record);
            this.title = metaTitle;
        }
    }
}


