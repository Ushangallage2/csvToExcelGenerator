
package com.example;

import javafx.application.Application;
import javafx.geometry.Insets;
import javafx.scene.Scene;
import javafx.scene.control.*;
import javafx.scene.layout.*;
import javafx.stage.FileChooser;
import javafx.stage.Stage;

import java.io.*;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.*;
import java.util.stream.Collectors;

import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVParser;
import org.apache.commons.csv.CSVRecord;
import org.apache.commons.io.input.BOMInputStream;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import javafx.concurrent.Task;

public class CSVProcessorApp extends Application {

    private File selectedCsvFile;
    private File processedExcelFile;
    private TextArea errorTextArea;
    private Label selectedFileLabel;
    private final CsvProcessor csvProcessor = new CsvProcessor();

    public static void main(String[] args) {
        launch(args);
    }

    @Override
    public void start(Stage primaryStage) {
        primaryStage.setTitle("CSV Processor");

        // UI Elements
        Button selectCsvButton = new Button("Select CSV File");
        selectedFileLabel = new Label("No file selected");
        Button processButton = new Button("Process CSV");
        Button viewExcelButton = new Button("View Excel File");
        Button saveExcelButton = new Button("Save Excel File");
        errorTextArea = new TextArea();
        errorTextArea.setEditable(false);
        errorTextArea.setPrefHeight(150);

        // Event Handlers
        selectCsvButton.setOnAction(e -> selectCsvFile());
        processButton.setOnAction(e -> processCsvFile());
        viewExcelButton.setOnAction(e -> viewExcelFile());
        saveExcelButton.setOnAction(e -> saveExcelFile());

        // Layout
        VBox layout = new VBox(10);
        layout.setPadding(new Insets(10));
        layout.getChildren().addAll(
                selectCsvButton,
                selectedFileLabel,
                processButton,
                viewExcelButton,
                saveExcelButton,
                new Label("Messages/Warnings:"),
                errorTextArea
        );

        Scene scene = new Scene(layout, 600, 500);
        primaryStage.setScene(scene);
        primaryStage.show();
    }

    private void selectCsvFile() {
        FileChooser fileChooser = new FileChooser();
        fileChooser.setTitle("Select CSV File");
        fileChooser.getExtensionFilters().add(new FileChooser.ExtensionFilter("CSV Files", "*.csv"));
        selectedCsvFile = fileChooser.showOpenDialog(new Stage());

        if (selectedCsvFile != null) {
            selectedFileLabel.setText("Selected file: " + selectedCsvFile.getAbsolutePath());
        } else {
            selectedFileLabel.setText("No file selected");
        }
    }

//    private void processCsvFile() {
//        if (selectedCsvFile == null) {
//            displayError("Please select a CSV file first.");
//            return;
//        }
//
//        // Use a Task to perform the CSV processing in a background thread
//        Task<Boolean> task = new Task<Boolean>() {
//            @Override
//            protected Boolean call() throws Exception {
//                String inputFilePath = selectedCsvFile.getAbsolutePath();
//                String outputFilePath = "output.xlsx"; // You can let the user choose this
//                try {
//                    // Process CSV file using your CsvProcessor
//                    boolean success = csvProcessor.processCsv(inputFilePath, outputFilePath, errorTextArea); // passing errorTextArea to the processCsv method
//                    if (success) {
//                        processedExcelFile = new File(outputFilePath);
//                    }
//                    return success;
//                } catch (Exception e) {
//                    updateMessage("Error processing CSV: " + e.getMessage());
//                    throw e;
//                }
//            }
//
//            @Override
//            protected void succeeded() {
//                super.succeeded();
//                boolean success = getValue();
//                if (success) {
//                    displayInfo("CSV Processing Complete! Excel file created at: output.xlsx");
//                } else {
//                    displayError("CSV Processing failed due to invalid headers or other errors.");
//                }
//            }
//
//            @Override
//            protected void failed() {
//                super.failed();
//                displayError(this.getMessage());
//            }
//        };
//
//        // Run the task in a background thread
//        new Thread(task).start();
//    }

    private void processCsvFile() {
        if (selectedCsvFile == null) {
            displayError("Please select a CSV file first.");
            return;
        }

        // Create and configure a progress indicator
        ProgressIndicator progressIndicator = new ProgressIndicator();
        progressIndicator.setProgress(-1.0); // Set to indeterminate mode
        progressIndicator.setVisible(true); // Make it visible while processing

        // Add the progress indicator to the layout
        VBox layout = (VBox) errorTextArea.getParent(); // Adjust the layout
        layout.getChildren().add(progressIndicator);

        // Set the output file path and check for the parent directory
        String outputFilePath = new File("output.xlsx").getAbsolutePath();
        File outputFile = new File(outputFilePath);

        // Get the parent directory of the output file
        File parentDir = outputFile.getParentFile();

        // Ensure the parent directory exists
        if (parentDir != null && !parentDir.exists()) {
            displayError("The output directory does not exist: " + outputFilePath);
            layout.getChildren().remove(progressIndicator); // Remove the indicator before returning
            return;
        }

        // Check available space on the drive where the output file will be saved
        long freeSpace = parentDir != null ? parentDir.getFreeSpace() : Long.MAX_VALUE; // Default to max space if no parent
        long requiredSpace = 1024 * 1024 * 5; // Estimate required space (5 MB for example)

        if (freeSpace < requiredSpace) {
            displayError("Not enough disk space to create the output file. Available space: " + freeSpace / (1024 * 1024) + " MB.");
            layout.getChildren().remove(progressIndicator); // Remove the indicator before returning
            return;
        }

        // Use a Task to perform the CSV processing in a background thread
        Task<Boolean> task = new Task<Boolean>() {
            @Override
            protected Boolean call() throws Exception {
                String inputFilePath = selectedCsvFile.getAbsolutePath();
                try {
                    boolean success = csvProcessor.processCsv(inputFilePath, outputFilePath, errorTextArea);
                    if (success) {
                        processedExcelFile = new File(outputFilePath);
                    }
                    return success;
                } catch (IOException e) {
                    updateMessage("Error processing CSV: " + e.getMessage());
                    javafx.application.Platform.runLater(() -> {
                        displayError(e.getMessage()); // Display user-friendly error
                    });
                    throw e; // Rethrow exception
                }
            }


            @Override
            protected void succeeded() {
                super.succeeded();
                progressIndicator.setVisible(false); // Hide the progress indicator when done
                boolean success = getValue();
                if (success) {
                    displayInfo("CSV Processing Complete! Excel file created at: output.xlsx");
                } else {
                    displayError("CSV Processing failed due to invalid headers or other errors.");
                }
                layout.getChildren().remove(progressIndicator); // Remove the indicator from layout
            }

            @Override
            protected void failed() {
                super.failed();
                progressIndicator.setVisible(false); // Hide the progress indicator on failure
                displayError(this.getMessage());
                layout.getChildren().remove(progressIndicator); // Remove the indicator from layout
            }
        };

        // Run the task in a background thread
        new Thread(task).start();
    }
//    private void processCsvFile() {
//        if (selectedCsvFile == null) {
//            displayError("Please select a CSV file first.");
//            return;
//        }
//
//        // Set the output file path and check for the parent directory
//        String outputFilePath = new File("output.xlsx").getAbsolutePath();
//        File outputFile = new File(outputFilePath);
//
//        // Get the parent directory of the output file
//        File parentDir = outputFile.getParentFile();
//
//        // Ensure the parent directory exists
//        if (parentDir != null && !parentDir.exists()) {
//            displayError("The output directory does not exist: " + outputFilePath);
//            return;
//        }
//
//        // Check available space on the drive where the output file will be saved
//        long freeSpace = parentDir != null ? parentDir.getFreeSpace() : Long.MAX_VALUE; // Default to max space if no parent
//        long requiredSpace = 1024 * 1024 * 5; // Estimate required space (5 MB for example)
//
//        if (freeSpace < requiredSpace) {
//            displayError("Not enough disk space to create the output file. Available space: " + freeSpace / (1024 * 1024) + " MB.");
//            return;
//        }
//
//        // Use a Task to perform the CSV processing in a background thread
//        Task<Boolean> task = new Task<Boolean>() {
//            @Override
//            protected Boolean call() throws Exception {
//                String inputFilePath = selectedCsvFile.getAbsolutePath();
//                try {
//                    // Process CSV file using your CsvProcessor
//                    boolean success = csvProcessor.processCsv(inputFilePath, outputFilePath, errorTextArea);
//                    if (success) {
//                        processedExcelFile = new File(outputFilePath);
//                    }
//                    return success;
//                } catch (Exception e) {
//                    updateMessage("Error processing CSV: " + e.getMessage());
//                    throw e;
//                }
//            }
//
//            @Override
//            protected void succeeded() {
//                super.succeeded();
//                boolean success = getValue();
//                if (success) {
//                    displayInfo("CSV Processing Complete! Excel file created at: output.xlsx");
//                } else {
//                    displayError("CSV Processing failed due to invalid headers or other errors.");
//                }
//            }
//
//            @Override
//            protected void failed() {
//                super.failed();
//                displayError(this.getMessage());
//            }
//        };
//
//        // Run the task in a background thread
//        new Thread(task).start();
//    }

    private void viewExcelFile() {
        if (processedExcelFile == null || !processedExcelFile.exists()) {
            displayError("Excel file not processed yet or does not exist.");
            return;
        }

        // Attempt to open the Excel file using the default system application based on OS type
        try {
            String os = System.getProperty("os.name").toLowerCase();
            if (os.contains("win")) {
                Runtime.getRuntime().exec(new String[]{"cmd", "/c", "start", processedExcelFile.getAbsolutePath()});
            } else if (os.contains("mac")) {
                Runtime.getRuntime().exec(new String[]{"/usr/bin/open", processedExcelFile.getAbsolutePath()});
            } else {
                displayError("Unsupported operating system for opening files.");
            }
        } catch (IOException e) {
            displayError("Error opening Excel file: " + e.getMessage());
        }
    }

    private void saveExcelFile() {
        if (processedExcelFile == null || !processedExcelFile.exists()) {
            displayError("Excel file not processed yet or does not exist.");
            return;
        }

        FileChooser fileChooser = new FileChooser();
        fileChooser.setTitle("Save Excel File");
        fileChooser.getExtensionFilters().add(new FileChooser.ExtensionFilter("Excel Files", "*.xlsx"));
        fileChooser.setInitialFileName("output.xlsx");

        File file = fileChooser.showSaveDialog(new Stage());
        if (file != null) {
            // Copy the processed file to the selected destination
            try {
                Files.copy(processedExcelFile.toPath(), file.toPath(), java.nio.file.StandardCopyOption.REPLACE_EXISTING);
                displayInfo("Excel file saved successfully to: " + file.getAbsolutePath());
            } catch (IOException e) {
                displayError("Error saving Excel file: " + e.getMessage());
            }
        }
    }

    private void displayError(String message) {
        errorTextArea.appendText("Error: " + message + "\n");
    }

    private void displayInfo(String message) {
        errorTextArea.appendText("Info: " + message + "\n");
    }

    // Inner class to encapsulate CSV processing logic
    public static class CsvProcessor {

        private static final String[] REQUIRED_HEADERS = {"Handle", "Title", "Product Category", "Option1 Name", "Option1 Value", "Option2 Name", "Option2 Value", "Variant SKU"};

        private static final Set<String> VALID_OPTION_TYPES = new HashSet<>(Arrays.asList("color", "colour", "size", "category", "group", "title"));

        public boolean processCsv(String inputFilePath, String outputFilePath, TextArea errorTextArea) throws IOException {
            try {
                if (!isFileWritable(outputFilePath)) {
                    javafx.application.Platform.runLater(() -> errorTextArea.appendText("Error: The output file '" + outputFilePath + "' is open or locked by another process. Please close it and try again.\n"));
                    return false;
                }

                Map<String, List<CSVRecord>> handleToRecordsMap = new HashMap<>();
                Map<String, List<ProductError>> errors = new HashMap<>();
                errors.put("Invalid - Duplicate SKUs", new ArrayList<>());
                errors.put("Invalid Options", new ArrayList<>());
                errors.put("Other Errors", new ArrayList<>());
                List<String> existingMetaProductHandles = new ArrayList<>();  // Track meta product handles

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
                            missingHeaders.add(requiredHeader);
                        }
                    }

                    if (!missingHeaders.isEmpty()) {
                        String errorMessage = "Warning: The following required headers are missing from your CSV file: " + missingHeaders +
                                ". Please update your CSV file headers to include: " + Arrays.toString(REQUIRED_HEADERS);
                        // Using Platform.runLater to update UI from background thread
                        javafx.application.Platform.runLater(() -> errorTextArea.appendText(errorMessage + "\n"));
                        return false;
                        // Returning false to indicate header validation failure
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


                        // Identify "Title/Default Title" Meta Products
                        List<CSVRecord> titleDefaultMetaProducts = records.stream()
                                .filter(r -> r.get("Option1 Name") != null && r.get("Option1 Name").equalsIgnoreCase("Title") &&
                                        r.get("Option1 Value") != null && r.get("Option1 Value").equalsIgnoreCase("Default Title"))
                                .collect(Collectors.toList());

                        // Enforce Single "Title/Default Title" Meta Product per Handle
                        if (titleDefaultMetaProducts.size() > 1) {
                            for (CSVRecord metaRecord : titleDefaultMetaProducts) {
                                errors.get("Other Errors").add(new ProductError(
                                        "Only one meta product with Option1 Name 'Title' and Option1 Value 'Default Title' is allowed per handle.", metaRecord));
                            }
                            continue; // Skip further processing for this handle
                        }

                        // Check if it is meta product (contains valid title)
                        List<CSVRecord> metaRecords = records.stream()
                                .filter(r -> r.get("Title") != null && !r.get("Title").isEmpty())
                                .collect(Collectors.toList());


                        // Enforce Single "Valid Title" Meta Product per Handle
                        if (metaRecords.size() > 1) {
                            for (CSVRecord metaRecord : metaRecords) {
                                errors.get("Other Errors").add(new ProductError(
                                        "Valid title option must have only one record: " + metaRecords.size() + " found.", metaRecord));
                            }
                            continue;
                        }

                        // If it's a  valid "Title/Default Title" Meta Product or "Valid Title" Meta Product, check for handle uniqueness
                        if (!titleDefaultMetaProducts.isEmpty() || !metaRecords.isEmpty()) {
                            if (existingMetaProductHandles.contains(handle)) {
                                for (CSVRecord metaRecord : titleDefaultMetaProducts.isEmpty() ? metaRecords : titleDefaultMetaProducts) {
                                    errors.get("Other Errors").add(new ProductError(
                                            "Meta product handle '" + handle + "' is not unique.", metaRecord));
                                }
                                continue; // Skip further processing for this handle
                            }
                            existingMetaProductHandles.add(handle);  // Add handle to the list
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
            return true;
        }

//        private static void writeErrorsToExcel(String outputFilePath, Map<String, List<ProductError>> errors) throws IOException {
//            try (Workbook workbook = new XSSFWorkbook()) {
//                for (Map.Entry<String, List<ProductError>> entry : errors.entrySet()) {
//                    writeErrorsToSheet(workbook, entry.getKey(), entry.getValue());
//                }
//                try (FileOutputStream outputStream = new FileOutputStream(outputFilePath)) {
//                    workbook.write(outputStream);
//                }
//            }
//        }

        private static void writeErrorsToExcel(String outputFilePath, Map<String, List<ProductError>> errors) throws IOException {
            try (Workbook workbook = new XSSFWorkbook()) {
                for (Map.Entry<String, List<ProductError>> entry : errors.entrySet()) {
                    writeErrorsToSheet(workbook, entry.getKey(), entry.getValue());
                }
                try (FileOutputStream outputStream = new FileOutputStream(outputFilePath)) {
                    workbook.write(outputStream);
                }
            } catch (FileNotFoundException e) {
                // Handle permission denied error specifically
                throw new IOException("Permission denied to write to: " + outputFilePath + ". Please ensure the file is not open in another application or adjust your file permissions.", e);
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

                // Write to the file
                try (FileOutputStream fileOut = new FileOutputStream(outputFilePath)) {
                    workbook.write(fileOut);
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
}


