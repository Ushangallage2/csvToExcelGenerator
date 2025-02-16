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
            List<CSVRecord> successfulRecords = new ArrayList<>();
            List<CSVRecord> imageEntries = new ArrayList<>(); // List to store image entries

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

                // Group records by handle AND REMOVE IMAGE ENTRIES
                List<CSVRecord> recordsToProcess = new ArrayList<>();
                for (CSVRecord record : parser) {
                    boolean isImageEntry = record.get("Option1 Name").isEmpty() &&
                            record.get("Option1 Value").isEmpty() &&
                            record.get("Option2 Name").isEmpty() &&
                            record.get("Option2 Value").isEmpty() &&
                            record.get("Variant SKU").isEmpty();

                    if (!isImageEntry) { // Only add non-image entries to the processing list
                        recordsToProcess.add(record);
                    } else {
                        System.out.println("Skipping image entry: " + record.get("Handle") + " - " + record.get("Variant SKU")); //optional logging
                    }
                }

                // Now process only the non-image entries
                for (CSVRecord record : recordsToProcess) {
                    String handle = record.get("Handle");
                    handleToRecordsMap.computeIfAbsent(handle, k -> new ArrayList<>()).add(record);
                }

                // Process each handle group
                for (Map.Entry<String, List<CSVRecord>> entry : handleToRecordsMap.entrySet()) {
                    String handle = entry.getKey();
                    List<CSVRecord> records = entry.getValue();

                    // Identify the meta product (record with a title)
                    List<CSVRecord> metaRecords = records.stream()
                            .filter(r -> r.get("Title") != null && !r.get("Title").isEmpty())
                            .collect(Collectors.toList());

                    // Validate meta product
                    if (metaRecords.size() > 1) {
                        for (CSVRecord metaRecord : metaRecords) {
                            errors.get("Other Errors").add(new ProductError(
                                    "Valid title option must have only one record: " + metaRecords.size() + " found.", metaRecord));
                        }
                        continue; // Skip this handle group if there are multiple meta products
                    }

                    CSVRecord metaRecord = metaRecords.isEmpty() ? null : metaRecords.get(0);

                    // Handle missing meta product
                    if (metaRecord == null) {
                        for (CSVRecord record : records) {
                            errors.get("Other Errors").add(new ProductError("Missing meta product title", record));
                        }
                        continue; // Skip this handle if no meta product found
                    }

                    String metaTitle = metaRecord.get("Title");
                    String metaOption1Name = metaRecord.get("Option1 Name");
                    String metaOption2Name = metaRecord.get("Option2 Name");

                    // Check for valid option names in the meta product
                    if (metaOption1Name.isEmpty() && metaOption2Name.isEmpty()) {
                        errors.get("Other Errors").add(new ProductError("Meta product must have at least one option name.", metaRecord));
                    }

                    // Check for duplicate option names
                    if (!metaOption1Name.isEmpty() && metaOption1Name.equalsIgnoreCase(metaOption2Name)) {
                        errors.get("Invalid Options").add(new ProductError("Duplicate option names found in meta product: " + metaOption1Name, metaRecord));
                    }

                    // Prepare to collect option names from the meta product
                    Set<String> optionNames = new HashSet<>();
                    if (!metaOption1Name.isEmpty()) {
                        optionNames.add(metaOption1Name.toLowerCase());
                    }
                    if (!metaOption2Name.isEmpty()) {
                        optionNames.add(metaOption2Name.toLowerCase());
                    }

                    for (CSVRecord record : records) {
                        String title = getCellValue(record, "Title");
                        String productCategory = getCellValue(record, "Product Category");
                        String sku = getCellValue(record, "Variant SKU");
                        String option1Name = getCellValue(record, "Option1 Name");
                        String option1Value = getCellValue(record, "Option1 Value");
                        String option2Name = getCellValue(record, "Option2 Name");
                        String option2Value = getCellValue(record, "Option2 Value");

                        boolean isValid = true; // Flag to track overall record validity

                        // Validate SKU
                        if (sku.isEmpty()) {
                            errors.get("Invalid - Duplicate SKUs").add(new ProductError("Missing SKU", record, metaTitle));
                            isValid = false;
                        } else if (skuSet.contains(sku)) {
                            errors.get("Invalid - Duplicate SKUs").add(new ProductError("Duplicate SKU found", record, metaTitle));
                            isValid = false;
                        } else {
                            skuSet.add(sku); // Add the SKU to the set if unique
                        }

                        // Inherit option names from meta product if missing
                        if (option1Name.isEmpty() && !metaOption1Name.isEmpty()) {
                            option1Name = metaOption1Name;
                        }
                        if (option2Name.isEmpty() && !metaOption2Name.isEmpty()) {
                            option2Name = metaOption2Name;
                        }

                        // If both option names are missing, flag as error
                        if (option1Name.isEmpty() && option2Name.isEmpty()) {
                            errors.get("Invalid Options").add(new ProductError(
                                    "Both Option1 Name and Option2 Name are missing for handle " + handle, record, metaTitle));
                            isValid = false;
                        }

                        // If valid, add to successful records
                        if (isValid && record.equals(metaRecord)) {
                            successfulRecords.add(record);
                        }
                    }
                }
            }

            // Log how many image entries were found
            System.out.println("Skipped image entries: " + imageEntries.size());

            // Write errors to Excel
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

        // Write count above header
        Row countRow = sheet.createRow(0);
        countRow.createCell(0).setCellValue("Count of " + sheetName + ": " + productErrors.size());

        // Create header row
        Row headerRow = sheet.createRow(1);
        String[] headers = {"Error Log", "Handle", "Title", "Product Category", "Option 1 Name", "Option 1 Value", "Option 2 Name", "Option 2 Value", "Variant SKU"};
        for (int i = 0; i < headers.length; i++) {
            headerRow.createCell(i).setCellValue(headers[i]);
        }

        // Write error data
        int rowNum = 2; // The next row to write errors to
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

    private static void writeSuccessfulRecordsToExcel(String outputFilePath, List<CSVRecord> successfulRecords) throws IOException {
        try (FileInputStream fileInputStream = new FileInputStream(outputFilePath);
             Workbook workbook = new XSSFWorkbook(fileInputStream)) {
            Sheet successSheet = workbook.createSheet("Success");

            // Write count above header
            Row countRow = successSheet.createRow(0);
            countRow.createCell(0).setCellValue("Count of Successful Records: " + successfulRecords.size());

            // Create header row for success
            Row headerRow = successSheet.createRow(1);
            String[] headers = {"Handle", "Title", "Product Category", "Option1 Name", "Option1 Value", "Option2 Name", "Option2 Value", "Variant SKU"};
            for (int i = 0; i < headers.length; i++) {
                headerRow.createCell(i).setCellValue(headers[i]);
            }

            // Write successful records
            int rowNum = 2; // Start writing successful records below the header
            for (CSVRecord record : successfulRecords) {
                Row row = successSheet.createRow(rowNum++);
                row.createCell(0).setCellValue(record.get("Handle"));
                row.createCell(1).setCellValue(record.get("Title"));
                row.createCell(2).setCellValue(record.get("Product Category"));
                row.createCell(3).setCellValue(record.get("Option1 Name"));
                row.createCell(4).setCellValue(record.get("Option1 Value"));
                row.createCell(5).setCellValue(record.get("Option2 Name"));
                row.createCell(6).setCellValue(record.get("Option2 Value"));
                row.createCell(7).setCellValue(record.get("Variant SKU"));
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

    private static String getCellValue(CSVRecord record, String header) {
        try {
            return record.get(header);
        } catch (IllegalArgumentException e) {
            System.err.println("Header missing: " + header);
            return ""; // Handle missing header gracefully
        }
    }

    private static class ProductError {
        String errorLog;
        String handle;
        String title;
        String productCategory;
        String option1Name;
        String option1Value;
        String option2Name;
        String option2Value;
        String variantSKU;

        public ProductError(String errorLog, CSVRecord record) {
            this.errorLog = errorLog;
            this.handle = getCellValue(record, "Handle");
            this.title = getCellValue(record, "Title"); // Get title from the record
            this.productCategory = getCellValue(record, "Product Category");
            this.option1Name = getCellValue(record, "Option1 Name");
            this.option1Value = getCellValue(record, "Option1 Value");
            this.option2Name = getCellValue(record, "Option2 Name");
            this.option2Value = getCellValue(record, "Option2 Value");
            this.variantSKU = getCellValue(record, "Variant SKU");
        }

        public ProductError(String errorLog, CSVRecord record, String title) {
            this.errorLog = errorLog;
            this.handle = getCellValue(record, "Handle");
            this.title = title; // Use provided title (meta title)
        }
    }
}