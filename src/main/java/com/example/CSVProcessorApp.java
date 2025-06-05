
package com.example;

import com.aspose.cells.SaveFormat;
import javafx.animation.FadeTransition;
import javafx.application.Application;
import javafx.fxml.FXML;
import javafx.geometry.Insets;
import javafx.geometry.Pos;
import javafx.scene.Node;
import javafx.scene.Scene;
import javafx.scene.control.*;
import javafx.scene.control.Button;
import javafx.scene.control.Dialog;
import javafx.scene.control.Label;
import javafx.scene.control.TextArea;
import javafx.scene.image.ImageView;
import javafx.scene.layout.*;
import javafx.scene.text.FontWeight;
import javafx.stage.FileChooser;
import javafx.stage.Stage;
import javafx.scene.image.Image;
import javafx.scene.text.Font;
import javafx.scene.text.FontWeight;
import javafx.scene.paint.Color;  // Import for Color
import javafx.scene.image.ImageView;

import java.awt.*;
import java.net.HttpURLConnection;
import java.nio.file.StandardCopyOption;
import java.util.List;
import java.util.prefs.Preferences;


import java.io.*;
import java.net.URL;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.*;
import java.util.stream.Collectors;

import javafx.util.Duration;
import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVParser;
import org.apache.commons.csv.CSVPrinter;
import org.apache.commons.csv.CSVRecord;
import org.apache.commons.io.input.BOMInputStream;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import javafx.concurrent.Task;
import javafx.concurrent.Task;  // Import Task
import javafx.application.Platform; // Import Platform

public class CSVProcessorApp extends Application {

    private File selectedCsvFile;
    private File processedExcelFile;
    private TextArea errorTextArea;
    private Label selectedFileLabel;
    @FXML
    private Label variationSelectedFileLabel;

    private final CsvProcessor csvProcessor = new CsvProcessor();
    private List<File> selectedCsvFiles = new ArrayList<>();
    private List<File> processedExcelFiles = new ArrayList<>();
    private List<File> savedExcelFiles = new ArrayList<>(); // List to track saved files
    private static final String RECENT_INPUT_FOLDER_KEY = "recent_input_folder";
    private static final String LAST_OUTPUT_FOLDER_KEY = "last_output_folder";
    private Preferences prefs = Preferences.userNodeForPackage(CSVProcessorApp.class);
    private String cssFilePath;
    private ProgressBar progressBar;  // Add ProgressBar
    private Label progressLabel; // Add Label for progress text


    private Button saveTemplate1;
    private Button processTemplate1;

    public static void main(String[] args) {
        launch(args);
    }


    private void deleteUnsavedOutputFiles() {
        for (File file : processedExcelFiles) {
            if (file.exists()) {
                boolean deleted = file.delete();
                if (deleted) {
                    displayInfo("Deleted temporary file: " + file.getName()); // Modified output
                } else {
                    displayError("Failed to delete file: " + file.getName());
                }
            }
        }
        processedExcelFiles.clear(); // Clear the list after deletion
    }

    private String loadPreference(String key, String defaultValue) {
        return prefs.get(key, defaultValue);
    }

    private void savePreference(String key, String value) {
        prefs.put(key, value);
    }


    // Method to get the default directory based on OS
    private String getDefaultDirectory() {
        String os = System.getProperty("os.name").toLowerCase();
        if (os.contains("win")) {
            return System.getProperty("user.home") + "\\Documents"; // Windows default
        } else {
            return System.getProperty("user.home") + "/Documents";  // macOS/Linux default
        }
    }

    private void applyDialogStyle(Dialog<?> dialog) {
        // Load CSS if available
        try {
            URL cssUrl = getClass().getClassLoader().getResource("Style.css");
            if (cssUrl != null) {
                dialog.getDialogPane().getStylesheets().add(cssUrl.toExternalForm());
            } else {
                System.err.println("CSS file not found. Styles will not be applied to dialog.");
            }
        } catch (Exception e) {
            System.err.println("Error loading CSS: " + e.getMessage());
        }


        // Set icon
        try {
            Image icon = new Image(Objects.requireNonNull(getClass().getClassLoader().getResourceAsStream("icon-excel.png")));
            Stage stage = (Stage) dialog.getDialogPane().getScene().getWindow();
            stage.getIcons().add(icon);
        } catch (Exception e) {
            System.err.println("Error loading icon: " + e.getMessage());
        }
    }

    private void clearSelectedFiles() {

        saveTemplate1.setText("save Template");
        processTemplate1.setText("process Template");
        errorTextArea.clear();
        processTemplate1.setText("process Template");
        processedExcelFiles.clear();
        selectedCsvFiles.clear();
        savedExcelFiles.clear();
        variationSelectedFileLabel.setText("No variant upload file selected");
        selectedFileLabel.setText("No CSV file selected");
    }



    @Override
    public void start(Stage primaryStage) {
        primaryStage.setTitle("Shopify CSV Fixer");

        String recentInputFolder = loadPreference(RECENT_INPUT_FOLDER_KEY, ""); // Default to empty
        String lastOutputFolder = loadPreference(LAST_OUTPUT_FOLDER_KEY, "");

        // Load the CSS file with error handling


// Load the CSS file with error handling
        String cssFilePath = null;
        try {
            URL cssUrl = getClass().getClassLoader().getResource("Style.css");
            if (cssUrl != null) {
                cssFilePath = cssUrl.toExternalForm();
                System.out.println("CSS file loaded from: " + cssFilePath);
            } else {
                System.err.println("CSS file not found. Styles will not be applied.");
                cssFilePath = null;
            }
        } catch (Exception e) {
            e.printStackTrace();
            cssFilePath = null;
        }
        // Load icon
        Image icon = new Image(Objects.requireNonNull(getClass().getClassLoader().getResourceAsStream("icon-excel.png")));
        primaryStage.getIcons().add(icon); // Set icon for the window


        // UI Elements
        Button selectCsvButton = new Button("Select CSV File");
        Button processButton = new Button("Process CSV");
        Button viewExcelButton = new Button("View Excel File");
        Button saveExcelButton = new Button("Save Excel File");
        Button clearSelectedFilesButton = new Button("Clear All");
        ComboBox<String> uploadFileTypeComboBox = new ComboBox<>();





        Button convertNumbersToCsvButton = new Button("Convert .numbers to .csv");

        // Add options to the ComboBox
        uploadFileTypeComboBox.getItems().addAll(
                "Variation Upload File",
                "Product Upload File"
        );

        // Optionally set a default value
        uploadFileTypeComboBox.setValue("Select Upload Template");


        processTemplate1 = new Button("process Template");



        saveTemplate1 = new Button("save Template");

        clearSelectedFilesButton.setOnAction(e -> {
            System.out.println("Clear all Button Clicked");
            clearSelectedFiles();
        });
        Button correctedOutputButton = new Button("Corrected Output");
        correctedOutputButton.setPrefWidth(250);
        correctedOutputButton.setOnAction(e -> {
            System.out.println("Corrected Output Button Clicked");
            chooseProcessedFileAndGenerateCorrectedOutput();
        });





        //eventlistener for the dropbox

        uploadFileTypeComboBox.setOnHidden(event -> {

            processTemplate1.setText("process Template");
            String selected = uploadFileTypeComboBox.getValue();
            Stage stage = (Stage) uploadFileTypeComboBox.getScene().getWindow();

            if ("Variation Upload File".equals(selected)) {
                processTemplate1.setText("process " + uploadFileTypeComboBox.getValue());
                handleProductUploadFile(stage);
            } else if ("Product Upload File".equals(selected)) {
                processTemplate1.setText("process " + uploadFileTypeComboBox.getValue());
                handleProductUploadFile(stage);
            }
        });




        processTemplate1.setOnAction(e -> {
            String selected = uploadFileTypeComboBox.getValue();

            // If the button is in "view product upload" mode, show files and return
            if ("view product upload".equalsIgnoreCase(processTemplate1.getText())) {
                viewExcelFile(); // Call  method to view processed files
                return;
            }

            try {
                if ("Variation Upload File".equals(selected)) {
                    processVariationUpload();
                    // You can add similar logic for variation upload if needed
                } else if ("Product Upload File".equals(selected)) {
                    processProductUpload();

                    // If processing is successful (processedExcelFiles is not empty and file exists)
                    if (!processedExcelFiles.isEmpty() && processedExcelFiles.get(0).exists()) {
                        processTemplate1.setText("view product upload");
                    }
                }
            } catch (IOException ex) {
                throw new RuntimeException("Error processing file: " + selected, ex);
            }
        });




        HBox buttonContainer = new HBox(10);
        buttonContainer.setAlignment(Pos.CENTER);
        buttonContainer.setPadding(new Insets(10));
        buttonContainer.getChildren().addAll(selectCsvButton, processButton, viewExcelButton, saveExcelButton);
        buttonContainer.getChildren().addAll(correctedOutputButton);

        HBox saveClearButtonContainer = new HBox(15);
        saveClearButtonContainer.setAlignment(Pos.CENTER_RIGHT);
        saveClearButtonContainer.setPadding(new Insets(10));

        // Create a spacer region
        Region spacer = new Region();
        HBox.setHgrow(spacer, Priority.ALWAYS);

        // Add nodes: inputTemplate1 | processTemplate1 | saveTemplate1 | spacer | clearSelectedFilesButton
        saveClearButtonContainer.getChildren().addAll(
                uploadFileTypeComboBox,
                processTemplate1,
                saveTemplate1,
                spacer,
                clearSelectedFilesButton
        );


        variationSelectedFileLabel= new Label("No Variation upload file selected");
        errorTextArea = new TextArea();
        errorTextArea.setEditable(false);
        errorTextArea.setPrefHeight(250);



        selectedFileLabel = new Label("No CSV file selected");
        errorTextArea = new TextArea();
        errorTextArea.setEditable(false);
        errorTextArea.setPrefHeight(150);

        // Instructions Label and Button
        Label instructionsLabel = new Label(
                "* You can input more than 1 CSV file\n" +
                        "* If you process one file over and over, output file will be named with the count of your attempt processed.\n" +
                        "* You can view an output file without saving it.\n" +
                        "* If you do not save an output file, it will be deleted upon closing the program."
        );
        instructionsLabel.setOpacity(0); // Initially invisible
        instructionsLabel.setVisible(false);

        FadeTransition ft = new FadeTransition(Duration.millis(1000), instructionsLabel);
        ft.setFromValue(0.0);
        ft.setToValue(1.0);


        Button toggleInstructionsButton = new Button("▼ Show Instructions");
        toggleInstructionsButton.setOnAction(e -> {
            boolean isVisible = instructionsLabel.isVisible();
            if (!isVisible) {
                instructionsLabel.setVisible(true);
                ft.playFromStart(); // Start fade-in animation

                // Get the preferred height of the instructionsLabel
                double instructionsHeight = instructionsLabel.getPrefHeight();

                // Current stage height
                double currentHeight = primaryStage.getHeight();

                // Total fixed height (assuming button height and some padding)
                double fixedHeight = 150 + 40; // Adjust as needed: 150 for TextArea and 40 for button, etc.

                // Calculate required height for the stage
                if (instructionsHeight > (currentHeight - fixedHeight)) {
                    primaryStage.setHeight(currentHeight + (instructionsHeight - (currentHeight - fixedHeight)) + 20); // Add some padding
                }

            } else {
                FadeTransition fadeOut = new FadeTransition(Duration.millis(500), instructionsLabel);
                fadeOut.setFromValue(1.0);
                fadeOut.setToValue(0.0);
                fadeOut.setOnFinished(event -> {
                    instructionsLabel.setVisible(false);
                    // No need to resize stage when hiding instructions
                });
                fadeOut.play();
            }
            toggleInstructionsButton.setText(isVisible ? "▼ Show Instructions" : "▲ Hide Instructions");
        });

        clearSelectedFilesButton.setStyle("-fx-background-color: #d4af37; -fx-text-fill: black; -fx-font-weight: bold; -fx-font-size: 14px; -fx-background-radius: 30; -fx-padding: 10 20; -fx-translate-x: -7;");
        uploadFileTypeComboBox.setStyle("-fx-background-color: #d4af37; -fx-text-fill: black; -fx-font-weight: bold; -fx-font-size: 14px; -fx-background-radius: 30; -fx-padding: 10 20; -fx-translate-x: -7;");

        // Bind buttons to stage width for responsiveness
        selectCsvButton.prefWidthProperty().bind(primaryStage.widthProperty().multiply(0.24));
        processButton.prefWidthProperty().bind(primaryStage.widthProperty().multiply(0.24));
        viewExcelButton.prefWidthProperty().bind(primaryStage.widthProperty().multiply(0.24));
        saveExcelButton.prefWidthProperty().bind(primaryStage.widthProperty().multiply(0.24));
        clearSelectedFilesButton.prefWidthProperty().bind(primaryStage.widthProperty().multiply(0.18));




        // Create the spacer
        Region spacer1 = new Region();
        HBox.setHgrow(spacer1, Priority.ALWAYS);

        // Wrap the button in an HBox and push it to the right
        HBox rightAlignedButtonBox = new HBox();
        rightAlignedButtonBox.getChildren().addAll(spacer1, convertNumbersToCsvButton);


        // Main Layout
        VBox layout = new VBox(10);
        layout.setPadding(new Insets(10));
        layout.getChildren().addAll(buttonContainer, saveClearButtonContainer, rightAlignedButtonBox, selectedFileLabel,variationSelectedFileLabel, new Label("Messages/Warnings:"), errorTextArea, toggleInstructionsButton, instructionsLabel);



        // Create and set the scene
        Scene scene = new Scene(layout, 950, 600);

        selectCsvButton.setOnAction(e -> {
            System.out.println("Select CSV Button Clicked");
            selectCsvFile();
        });
        processButton.setOnAction(e -> {
            System.out.println("Process Button Clicked");
            processCsvFile();
        });
        viewExcelButton.setOnAction(e -> {
            System.out.println("View Excel Button Clicked");
            viewExcelFile();
        });
        saveExcelButton.setOnAction(e -> {
            System.out.println("Save Excel Button Clicked");
            saveExcelFile();
        });


        saveTemplate1.setOnAction(e -> {
            System.out.println("Save Excel Button Clicked");
            saveExcelFile();
        });


        convertNumbersToCsvButton.setOnAction(e -> {
            FileChooser fileChooser = new FileChooser();
            fileChooser.setTitle("Select .numbers File");
            fileChooser.getExtensionFilters().add(new FileChooser.ExtensionFilter("Numbers Files (*.numbers)", "*.numbers"));
            File numbersFile = fileChooser.showOpenDialog(primaryStage);
            if (numbersFile == null) {
                displayError("No .numbers file selected.");
                return;
            }

            FileChooser saveChooser = new FileChooser();
            saveChooser.setTitle("Save Converted CSV");
            saveChooser.getExtensionFilters().add(new FileChooser.ExtensionFilter("CSV Files (*.csv)", "*.csv"));
            File csvFile = saveChooser.showSaveDialog(primaryStage);
            if (csvFile == null) {
                displayError("No save location selected.");
                return;
            }

            try {
                convertNumbersToCsv(numbersFile, csvFile); // This uses Aspose.Cells
                displayInfo("Conversion successful! Saved to: " + csvFile.getAbsolutePath());
            } catch (NullPointerException npe) {
                npe.printStackTrace();
                displayError(
                        "Conversion failed: This .numbers file cannot be opened by Aspose.Cells. " +
                                "Most modern Apple Numbers files are not supported. " +
                                "Please export your file as CSV from Apple Numbers on a Mac, or use an online converter such as https://cloudconvert.com/numbers-to-csv."
                );
            } catch (Exception ex) {
                ex.printStackTrace(); // This will print the full error in your console
                displayError("Conversion failed: " + (ex.getMessage() != null ? ex.getMessage() : ex.toString()));
            }


        });









        if (cssFilePath != null) {
            scene.getStylesheets().add(cssFilePath); // Apply the CSS
        } else {
            System.err.println("CSS file not loaded. Styles will not be applied.");
        }

        // Set the scene to the primaryStage
        primaryStage.setScene(scene);

        // Add a listener to the width property
        primaryStage.widthProperty().addListener((observable, oldValue, newValue) -> {
            // Here you can adjust components if needed
            // For example, you could change margins or update layout properties
            System.out.println("Width changed from " + oldValue + " to " + newValue);

            // You could add any responsive behavior you may need here
            // For example, you might want to adjust sizes of other controls
            layout.setPrefWidth(newValue.doubleValue()); // Example of applying width to the layout
        });





        if (!recentInputFolder.isEmpty()) {
            File initialDir = new File(recentInputFolder);
            if (initialDir.exists()) {
                System.out.println("Loading recent input folder: " + recentInputFolder);
            }
        }




        primaryStage.setOnCloseRequest(event -> {
            List<File> unsavedFiles = new ArrayList<>(processedExcelFiles);
            unsavedFiles.removeAll(savedExcelFiles); // Files processed but not saved

            if (!unsavedFiles.isEmpty()) {
                Alert alert = new Alert(Alert.AlertType.CONFIRMATION);
                applyDialogStyle(alert);
                String cssFilePathLocal = getClass().getClassLoader().getResource("Style.css").toExternalForm();

                // Apply the custom CSS for the alert
                if (cssFilePathLocal != null) {
                    alert.getDialogPane().getStylesheets().add(cssFilePathLocal);
                }

                alert.setTitle("Unsaved Files");
                alert.setHeaderText("You have unsaved processed files.");

                StringBuilder sb = new StringBuilder();
                sb.append("The following files have been processed but not saved:\n");
                for (File file : unsavedFiles) {
                    sb.append("- ").append(file.getName()).append("\n");
                }
                sb.append("Please save these files if you wish to keep them.  Temporary files will be deleted on exit."); // more clear message

                alert.setContentText(sb.toString());

                // Load the custom icon for the alert
                Image alertIcon = null;
                try {
                    InputStream alertIconStream = getClass().getClassLoader().getResourceAsStream("icon-excel.png");
                    if (alertIconStream != null) {
                        alertIcon = new Image(alertIconStream);

                        // Set the icon for the alert dialog's title bar
                        Stage alertStage = (Stage) alert.getDialogPane().getScene().getWindow();
                        if (alertStage != null) { // Defensive check
                            alertStage.getIcons().clear();  // Remove default icon
                            alertStage.getIcons().add(alertIcon);
                        }
                    } else {
                        displayError("Icon file not found: icon-excel.png");
                    }
                } catch (Exception e) {
                    displayError("Error loading alert icon: " + e.getMessage());
                }

                ImageView iconView = new ImageView(alertIcon); // Display icon next to content
                iconView.setFitWidth(20); // Set appropriate size
                iconView.setFitHeight(20);

                // Create a toolbar-like layout to align icon with the title
                HBox header = new HBox();
                header.setAlignment(Pos.CENTER_LEFT); // Align left

                // Add a label for the title
                Label titleLabel = new Label("  Unsaved Files");
                titleLabel.setTextFill(Color.BLACK); // Ensure the title is black
                titleLabel.setFont(Font.font("Arial", FontWeight.BOLD, 14)); // Set font style if needed

                header.getChildren().add(titleLabel); // Add title next to the icon

                // Set the custom header
                alert.getDialogPane().setHeader(header); // Assign custom header to alert pane
                Optional<ButtonType> result = alert.showAndWait();


                // Check the result of the alert
                if (result.isPresent() && result.get() == ButtonType.CANCEL) {
                    event.consume(); // Consume the close event to prevent closing
                    return; // Exit the event handler
                }
            }

            // Delete all processed files, regardless of whether they were saved.
            deleteUnsavedOutputFiles();  // ALWAYS delete the temporary files

        });

        primaryStage.show();


    }

    public void convertNumbersToCsv(File numbersFile, File csvFile) throws Exception {
        // Do NOT catch exceptions here!
        com.aspose.cells.Workbook workbook = new com.aspose.cells.Workbook(numbersFile.getAbsolutePath());
        workbook.save(csvFile.getAbsolutePath(), com.aspose.cells.SaveFormat.CSV);
    }




    private String parseFileUrlFromJson(String json) {
        // Very basic parsing (production code should use a JSON library)
        int urlIndex = json.indexOf("\"Url\":\"");
        if (urlIndex == -1) return null;
        int start = urlIndex + 7;
        int end = json.indexOf("\"", start);
        if (end == -1) return null;
        return json.substring(start, end).replace("\\/", "/");
    }




    private void processProductUpload() {
        errorTextArea.clear();

        if (selectedCsvFiles.isEmpty()) {
            errorTextArea.setText("No file selected.");
            return;
        }

        File csvFile = selectedCsvFiles.get(0);
        List<String> errors = new ArrayList<>();

        // Only these columns will be exported
        List<String> exportHeaders = Arrays.asList("variation_name", "option1", "option2", "product_code");

        try {
            List<String> lines = Files.readAllLines(csvFile.toPath(), StandardCharsets.UTF_8);
            String cleanedHeaderLine = lines.get(0).replaceAll(",\\s*$", "");
            lines.set(0, cleanedHeaderLine);
            String cleanedCsvContent = String.join("\n", lines);

            try (Reader cleanedReader = new StringReader(cleanedCsvContent)) {
                CSVParser parser = new CSVParser(cleanedReader, CSVFormat.DEFAULT
                        .withFirstRecordAsHeader()
                        .withIgnoreHeaderCase()
                        .withTrim());

                Map<String, Integer> headerMap = parser.getHeaderMap();
                for (String required : exportHeaders) {
                    if (!headerMap.containsKey(required)) {
                        errorTextArea.setText("Missing required header: " + required);
                        return;
                    }
                }

                // 1. Group records by variation_name
                List<List<CSVRecord>> allGroups = new ArrayList<>();
                List<String> groupNames = new ArrayList<>();
                List<CSVRecord> currentGroup = new ArrayList<>();
                for (CSVRecord record : parser) {
                    // Check for blank record
                    boolean isBlankRecord = true;
                    for (String value : exportHeaders) {
                        if (!record.get(value).trim().isEmpty()) {
                            isBlankRecord = false;
                            break;
                        }
                    }
                    if (isBlankRecord) continue;

                    String variationName = record.get("variation_name").trim();
                    if (!variationName.isEmpty()) {
                        if (!currentGroup.isEmpty()) {
                            allGroups.add(new ArrayList<>(currentGroup));
                            currentGroup.clear();
                        }
                        groupNames.add(variationName);
                    }
                    currentGroup.add(record);
                }
                if (!currentGroup.isEmpty()) {
                    allGroups.add(currentGroup);
                }

                // 2. Check for duplicate variation_name groups
                Map<String, List<Integer>> variationNameToGroups = new HashMap<>();
                for (int i = 0; i < groupNames.size(); i++) {
                    String name = groupNames.get(i);
                    variationNameToGroups.computeIfAbsent(name, k -> new ArrayList<>()).add(i);
                }

                // 3. Map product_code to groups
                Map<String, List<Integer>> productCodeToGroups = new HashMap<>();
                for (int i = 0; i < allGroups.size(); i++) {
                    for (CSVRecord record : allGroups.get(i)) {
                        String productCode = record.get("product_code").trim();
                        if (!productCode.isEmpty()) {
                            productCodeToGroups.computeIfAbsent(productCode, k -> new ArrayList<>()).add(i);
                        }
                    }
                }

                // 4. Identify invalid groups (duplicate variation_name, duplicate product_code, missing fields)
                Set<Integer> invalidGroupIndexes = new HashSet<>();
                // a) Duplicate variation_name
                for (Map.Entry<String, List<Integer>> entry : variationNameToGroups.entrySet()) {
                    if (entry.getValue().size() > 1) {
                        invalidGroupIndexes.addAll(entry.getValue());
                    }
                }
                // b) Duplicate product_code across groups
                for (Map.Entry<String, List<Integer>> entry : productCodeToGroups.entrySet()) {
                    if (entry.getValue().size() > 1) {
                        invalidGroupIndexes.addAll(entry.getValue());
                    }
                }
                // c) Validation for each group (option1 required for header, product_code required/unique in group)
                for (int i = 0; i < allGroups.size(); i++) {
                    List<CSVRecord> group = allGroups.get(i);
                    Set<String> localProductCodes = new HashSet<>();
                    for (int j = 0; j < group.size(); j++) {
                        CSVRecord record = group.get(j);
                        String productCode = record.get("product_code").trim();
                        String option1 = record.get("option1").trim();
                        String variationName = record.get("variation_name").trim();

                        if (j == 0 && option1.isEmpty()) {
                            invalidGroupIndexes.add(i);
                        }
                        if (productCode.isEmpty() || !localProductCodes.add(productCode)) {
                            invalidGroupIndexes.add(i);
                        }
                    }
                }

                // 5. Separate valid/invalid groups
                List<List<CSVRecord>> validGroups = new ArrayList<>();
                List<List<CSVRecord>> invalidGroups = new ArrayList<>();
                for (int i = 0; i < allGroups.size(); i++) {
                    if (invalidGroupIndexes.contains(i)) {
                        invalidGroups.add(allGroups.get(i));
                    } else {
                        validGroups.add(allGroups.get(i));
                    }
                }

                // 6. Write valid groups
                if (!validGroups.isEmpty()) {
                    File outFile = new File(csvFile.getParent(), "product_upload_processed.csv");
                    try (BufferedWriter writer = Files.newBufferedWriter(outFile.toPath(), StandardCharsets.UTF_8)) {
                        CSVPrinter printer = new CSVPrinter(writer, CSVFormat.DEFAULT.withHeader(exportHeaders.toArray(new String[0])));
                        for (List<CSVRecord> group : validGroups) {
                            for (CSVRecord rec : group) {
                                List<String> row = exportHeaders.stream()
                                        .map(h -> rec.get(h))
                                        .collect(Collectors.toList());
                                printer.printRecord(row);
                            }
                        }
                        printer.flush();
                        processedExcelFiles.clear();
                        processedExcelFiles.add(outFile);
                        errors.add("Processed file Temporary saved as: " + outFile.getAbsolutePath());
                        saveTemplate1.setText("save Product Upload File");
                    }
                } else {
                    errors.add("No valid product groups to write.");
                }

                // 7. Write invalid groups
                if (!invalidGroups.isEmpty()) {
                    File invalidFile = new File(csvFile.getParent(), "invalid.csv");
                    try (BufferedWriter writer = Files.newBufferedWriter(invalidFile.toPath(), StandardCharsets.UTF_8)) {
                        CSVPrinter printer = new CSVPrinter(writer, CSVFormat.DEFAULT.withHeader(exportHeaders.toArray(new String[0])));
                        for (List<CSVRecord> group : invalidGroups) {
                            for (CSVRecord rec : group) {
                                List<String> row = exportHeaders.stream()
                                        .map(h -> rec.get(h))
                                        .collect(Collectors.toList());
                                printer.printRecord(row);
                            }
                        }
                        printer.flush();
                        processedExcelFiles.add(invalidFile);
                        errors.add("Invalid records written to: " + invalidFile.getAbsolutePath());
                    }
                }

                errorTextArea.setText(String.join("\n", errors));
            }
        } catch (Exception ex) {
            errorTextArea.setText("Error processing file: " + ex.getMessage());
            ex.printStackTrace();
        }
    }




    private void processVariationUpload() throws IOException {
        errorTextArea.clear();

        if (selectedCsvFiles.isEmpty()) {
            errorTextArea.setText("No file selected.");
            return;
        }

        File csvFile = selectedCsvFiles.get(0);
        List<String> errors = new ArrayList<>();
        List<VariationRecord> validRecords = new ArrayList<>();
        Set<String> variationNames = new HashSet<>();
        Set<String> productCodes = new HashSet<>();
        Set<String> duplicateVariationNames = new HashSet<>();
        Set<String> duplicateProductCodes = new HashSet<>();

        // Step 1: Read raw lines
        List<String> lines = Files.readAllLines(csvFile.toPath(), StandardCharsets.UTF_8);

        // Step 2: Trim trailing commas in header
        String cleanedHeaderLine = lines.get(0).replaceAll(",\\s*$", "");

        // Step 3: Replace header and rejoin for parser
        lines.set(0, cleanedHeaderLine);
        String cleanedCsvContent = String.join("\n", lines);

        // Step 4: Parse cleaned content
        try (Reader cleanedReader = new StringReader(cleanedCsvContent)) {
            CSVParser parser = new CSVParser(cleanedReader, CSVFormat.DEFAULT
                    .withFirstRecordAsHeader()
                    .withIgnoreHeaderCase()
                    .withTrim());


            Map<String, Integer> headerMap = parser.getHeaderMap();
            List<String> requiredHeaders = Arrays.asList("variation_name", "option1", "option2", "product_code");

            // Check required headers
            for (String required : requiredHeaders) {
                if (!headerMap.containsKey(required)) {
                    errorTextArea.setText("Missing required header: " + required);
                    return;
                }
            }

            int rowNum = 1 + 1; // header + 1-based indexing
            for (CSVRecord record : parser) {
                String variationName = record.get("variation_name").trim();
                String option1 = record.get("option1").trim();
                String option2 = record.get("option2").trim();
                String productCode = record.get("product_code").trim();

                if (variationName.isEmpty()) {
                    rowNum++;
                    continue;
                }

                if (option1.isEmpty() && option2.isEmpty()) {
                    errors.add("Row " + rowNum + ": Must have at least option1 or option2 for variation_name: " + variationName);
                    rowNum++;
                    continue;
                }

                if (!option2.isEmpty() && option1.isEmpty()) {
                    errors.add("Row " + rowNum + ": Has option2 but missing option1 for variation_name: " + variationName);
                    rowNum++;
                    continue;
                }

                if (!variationNames.add(variationName)) {
                    duplicateVariationNames.add(variationName);
                    rowNum++;
                    continue;
                }

                if (!productCode.isEmpty() && !productCodes.add(productCode)) {
                    duplicateProductCodes.add(productCode);
                    rowNum++;
                    continue;
                }

                validRecords.add(new VariationRecord(variationName, option1, option2, productCode));
                rowNum++;
            }

            if (!duplicateVariationNames.isEmpty()) {
                errors.add("Duplicate variation_name(s): " + String.join(", ", duplicateVariationNames));
            }
            if (!duplicateProductCodes.isEmpty()) {
                errors.add("Duplicate product_code(s): " + String.join(", ", duplicateProductCodes));
            }

            if (!validRecords.isEmpty()) {
                File outFile = new File(csvFile.getParent(), "variation_upload_processed.csv");
                try (BufferedWriter writer = Files.newBufferedWriter(outFile.toPath(), StandardCharsets.UTF_8)) {
                    CSVPrinter printer = new CSVPrinter(writer, CSVFormat.DEFAULT
                            .withHeader("variation_name", "option1", "option2", "meta_product_code"));

                    for (VariationRecord record : validRecords) {
                        printer.printRecord(
                                record.getVariationName(),
                                record.getOption1(),
                                record.getOption2(),
                                record.getProductCode()
                        );
                    }

                    printer.flush();
                    processedExcelFiles.clear();  // keep only the processedVariantupload files
                    processedExcelFiles.add(outFile);
                    errors.add("Processed file Temporary saved as: " + outFile.getAbsolutePath());
                    saveTemplate1.setText("save Variation Upload File");
                }
            } else {
                errors.add("No valid records to write.");
            }

            errorTextArea.setText(String.join("\n", errors));

        } catch (Exception ex) {
            errorTextArea.setText("Error processing file: " + ex.getMessage());
            ex.printStackTrace();
        }
    }




    private String getCellValue(String[] row, int index) {
        return (index >= 0 && index < row.length) ? row[index].trim() : "";
    }









    private void handleProductUploadFile(Stage stage) {
        // Your existing code, but replace 'yourStage' with this 'stage' parameter

        saveTemplate1.setText("save Template");

        FileChooser fileChooser = new FileChooser();
        fileChooser.setTitle("Select CSV or Numbers Files");

        FileChooser.ExtensionFilter csvFilter = new FileChooser.ExtensionFilter("CSV Files (*.csv)", "*.csv");
        FileChooser.ExtensionFilter numbersFilter = new FileChooser.ExtensionFilter("Numbers Files (*.numbers)", "*.numbers");
        fileChooser.getExtensionFilters().addAll(csvFilter, numbersFilter);

        // Load preference and set initial directory as before
        String recentInputFolder = loadPreference(RECENT_INPUT_FOLDER_KEY, "");
        if (!recentInputFolder.isEmpty()) {
            File initialDir = new File(recentInputFolder);
            if (initialDir.exists() && initialDir.isDirectory()) {
                fileChooser.setInitialDirectory(initialDir);
            }
        }

        List<File> selectedFiles = fileChooser.showOpenMultipleDialog(stage);

        if (selectedFiles == null || selectedFiles.isEmpty()) {
            displayError("No files selected. Please select a file.");
            return;
        }

        savePreference(RECENT_INPUT_FOLDER_KEY, selectedFiles.get(0).getParent());

        List<File> csvFilesToProcess = new ArrayList<>();

        for (File file : selectedFiles) {
            String fileName = file.getName().toLowerCase();

            if (fileName.endsWith(".csv")) {
                csvFilesToProcess.add(file);
            } else if (fileName.endsWith(".numbers")) {
                displayError(
                        "The selected file \"" + file.getName() + "\" is a .numbers file.\n\n" +
                                "Please convert it to CSV using this online tool:\n" +
                                "https://cloudconvert.com/numbers-to-csv"
                );
            } else {
                displayError("Invalid file type selected: " + file.getName() + "\nPlease select only .csv or .numbers files.");
            }
        }

        if (!csvFilesToProcess.isEmpty()) {
            selectedCsvFiles.addAll(csvFilesToProcess);

            // Update label with count
            int selectedCount = csvFilesToProcess.size();
            variationSelectedFileLabel.setText(selectedCount + (selectedCount == 1 ? " file selected." : " files selected."));

            System.out.println("CSV files ready for processing:");
            csvFilesToProcess.forEach(f -> System.out.println(f.getAbsolutePath()));

            // TODO: Add your CSV processing logic here
        } else {
            variationSelectedFileLabel.setText("No valid CSV files selected.");
        }





    }


    //questionable select file method

    private void handleVariationUploadFile(Stage stage) {
        // Your existing code, but replace 'yourStage' with this 'stage' parameter
        FileChooser fileChooser = new FileChooser();
        fileChooser.setTitle("Select CSV or Numbers Files");

        FileChooser.ExtensionFilter csvFilter = new FileChooser.ExtensionFilter("CSV Files (*.csv)", "*.csv");
        FileChooser.ExtensionFilter numbersFilter = new FileChooser.ExtensionFilter("Numbers Files (*.numbers)", "*.numbers");
        fileChooser.getExtensionFilters().addAll(csvFilter, numbersFilter);

        // Load preference and set initial directory as before
        String recentInputFolder = loadPreference(RECENT_INPUT_FOLDER_KEY, "");
        if (!recentInputFolder.isEmpty()) {
            File initialDir = new File(recentInputFolder);
            if (initialDir.exists() && initialDir.isDirectory()) {
                fileChooser.setInitialDirectory(initialDir);
            }
        }

        List<File> selectedFiles = fileChooser.showOpenMultipleDialog(stage);

        if (selectedFiles == null || selectedFiles.isEmpty()) {
            displayError("No files selected. Please select a file.");
            return;
        }

        savePreference(RECENT_INPUT_FOLDER_KEY, selectedFiles.get(0).getParent());

        List<File> csvFilesToProcess = new ArrayList<>();

        for (File file : selectedFiles) {
            String fileName = file.getName().toLowerCase();

            if (fileName.endsWith(".csv")) {
                csvFilesToProcess.add(file);
            } else if (fileName.endsWith(".numbers")) {
                displayError(
                        "The selected file \"" + file.getName() + "\" is a .numbers file.\n\n" +
                                "Please convert it to CSV using this online tool:\n" +
                                "https://cloudconvert.com/numbers-to-csv"
                );
            } else {
                displayError("Invalid file type selected: " + file.getName() + "\nPlease select only .csv or .numbers files.");
            }
        }

        if (!csvFilesToProcess.isEmpty()) {
            // Update label with count
            int selectedCount = csvFilesToProcess.size();
            variationSelectedFileLabel.setText(selectedCount + (selectedCount == 1 ? " file selected." : " files selected."));

            System.out.println("CSV files ready for processing:");
            csvFilesToProcess.forEach(f -> System.out.println(f.getAbsolutePath()));

            // TODO: Add your CSV processing logic here
        } else {
            selectedFileLabel.setText("No valid CSV files selected.");
        }

    }



    private void selectCsvFile() {
        FileChooser fileChooser = new FileChooser();
        fileChooser.setTitle("Select CSV Files");
        fileChooser.getExtensionFilters().add(new FileChooser.ExtensionFilter("CSV Files", "*.csv"));

        // Set the initial directory from preferences
        String recentInputFolder = loadPreference(RECENT_INPUT_FOLDER_KEY, "");
        if (!recentInputFolder.isEmpty()) {
            fileChooser.setInitialDirectory(new File(recentInputFolder));
        }

        List<File> selectedFiles = fileChooser.showOpenMultipleDialog(new Stage());

        if (selectedFiles != null && !selectedFiles.isEmpty()) {
            selectedCsvFiles.clear(); // Clear previous selections
            selectedCsvFiles.addAll(selectedFiles);
            selectedFileLabel.setText(selectedFiles.size() + " files selected.");

            // Save the recent input folder path
            savePreference(RECENT_INPUT_FOLDER_KEY, selectedFiles.get(0).getParent());
        } else {
            selectedCsvFiles.clear();
            selectedFileLabel.setText("No files selected.");
        }
    }




    // Helper method to find the header row dynamically
    private int findHeaderRow(Sheet sheet, String[] expectedHeaders) {
        for (Row row : sheet) {
            boolean isHeaderRow = true;
            for (int i = 0; i < expectedHeaders.length; i++) {
                Cell cell = row.getCell(i);
                if (cell == null || !cell.getStringCellValue().equals(expectedHeaders[i])) {
                    isHeaderRow = false;
                    break;
                }
            }
            if (isHeaderRow) {
                return row.getRowNum(); // Return the index of the header row
            }
        }
        return -1; // Header row not found
    }

    private void processCsvFile() {
        if (selectedCsvFiles.isEmpty()) {
            displayError("Please select CSV files first.");
            return;
        }

        VBox layout = (VBox) errorTextArea.getParent();
        ProgressIndicator progressIndicator = new ProgressIndicator();
        progressIndicator.setProgress(-1.0);
        progressIndicator.setVisible(true);
        layout.getChildren().add(progressIndicator);

        Task<Void> processingTask = new Task<Void>() {
            @Override
            protected Void call() throws Exception {
                int totalFiles = selectedCsvFiles.size();
                for (int i = 0; i < totalFiles; i++) {
                    File csvFile = selectedCsvFiles.get(i);
                    String inputFilePath = csvFile.getAbsolutePath();
                    String baseName = csvFile.getName().replaceFirst("[.][^.]+$", ""); // Filename without extension

                    int attemptCount = getAttemptCount(baseName);
                    String outputFilePath = baseName + "_attempt_" + attemptCount + ".xlsx"; // Unique name

                    try {
                        boolean success = csvProcessor.processCsv(inputFilePath, outputFilePath, errorTextArea);
                        if (success) {
                            File outputFile = new File(outputFilePath);
                            processedExcelFiles.add(outputFile);

                            // Check if there are any errors in the output file
                            boolean hasErrors = csvProcessor.hasErrors(outputFile);
                            if (hasErrors) {
                                Platform.runLater(() -> displayError("There are errors in this file. Please check: " + csvFile.getName() + " 😥"));
                            } else {
                                Platform.runLater(() -> displayInfo("There is no error in the file! " + csvFile.getName() + "😊"));
                            }
                        }
                    } catch (IOException e) {
                        Platform.runLater(() -> displayError("Error processing " + csvFile.getName() + ": " + e.getMessage()));
                    }
                                // Provide some delay to allow visibility of processing feedback, if needed.
                                Thread.sleep(100); //slight delay
                            }

                            return null;
                        }

            @Override
            protected void succeeded() {
                super.succeeded();
                Platform.runLater(() -> {
                    displayInfo("All files processed successfully.");
                    layout.getChildren().remove(progressIndicator); // Remove ProgressIndicator
                });
            }

            @Override
            protected void failed() {
                super.failed();
                Throwable error = getException(); // Get the actual exception
                Platform.runLater(() -> {
                    displayError("File processing failed: " + (error != null ? error.getMessage() : "Unknown error"));
                    layout.getChildren().remove(progressIndicator); // Remove ProgressIndicator
                });
            }
        };

        new Thread(processingTask).start(); // Start the processing task in a new thread
    }


    private void chooseProcessedFileAndGenerateCorrectedOutput() {
        if (processedExcelFiles.isEmpty()) {
            displayError("No processed Excel files available.");
            return;
        }

        ChoiceDialog<File> dialog = new ChoiceDialog<>(processedExcelFiles.get(0), processedExcelFiles);
        dialog.setTitle("Choose Processed File");
        dialog.setHeaderText("Select the processed Excel file to generate corrected output from:");
        dialog.setContentText("Choose a file:");

        applyDialogStyle(dialog);

        Optional<File> result = dialog.showAndWait();
        result.ifPresent(this::generateCorrectedOutput);
    }



    private void generateCorrectedOutput(File selectedExcelFile) {
        if (selectedExcelFile == null) {
            displayError("No file selected.");
            return;
        }

        try (FileInputStream fis = new FileInputStream(selectedExcelFile);
             Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet successSheet = workbook.getSheet("Success");
            if (successSheet == null) {
                displayError("Sheet 'Success' not found in the selected Excel file.");
                return;
            }

            Workbook newWorkbook = new XSSFWorkbook();
            Sheet newSheet = newWorkbook.createSheet("Filtered Success");

            Map<String, Integer> columnIndexMap = new HashMap<>();
            columnIndexMap.put("Handle", 0);
            columnIndexMap.put("Title", 1);
            columnIndexMap.put("Option1 Name", 8);
            columnIndexMap.put("Option1 Value", 9);
            columnIndexMap.put("Option2 Name", 11);
            columnIndexMap.put("Option2 Value", 12);
            columnIndexMap.put("Variant SKU", 17);

            Row headerRow = newSheet.createRow(0);
            for (Map.Entry<String, Integer> entry : columnIndexMap.entrySet()) {
                Cell cell = headerRow.createCell(entry.getValue());
                cell.setCellValue(entry.getKey());
            }

            Map<String, Integer> sourceColumnIndexMap = new HashMap<>();
            Row headerRowSource = successSheet.getRow(1);
            if (headerRowSource != null) {
                for (int i = 0; i < headerRowSource.getLastCellNum(); i++) {
                    Cell cell = headerRowSource.getCell(i);
                    if (cell != null) {
                        String cellValue = cell.getStringCellValue().trim();
                        sourceColumnIndexMap.put(cellValue, i);
                    }
                }
            }

            Set<String> requiredSourceColumns = new HashSet<>(columnIndexMap.keySet());
            requiredSourceColumns.add("Meta Status");
            if (!sourceColumnIndexMap.keySet().containsAll(requiredSourceColumns)) {
                requiredSourceColumns.removeAll(sourceColumnIndexMap.keySet());
                displayError("Missing required columns in 'Success' sheet: " + String.join(", ", requiredSourceColumns));
                return;
            }

            List<String> variantSKUs = new ArrayList<>();

            int rowIndex = 1;
            for (int i = 2; i <= successSheet.getLastRowNum(); i++) {
                Row dataRow = successSheet.getRow(i);
                if (dataRow != null) {
                    Integer metaStatusColumnIndex = sourceColumnIndexMap.get("Meta Status");
                    if (metaStatusColumnIndex != null) {
                        Cell metaStatusCell = dataRow.getCell(metaStatusColumnIndex);
                        String metaStatus = (metaStatusCell != null && metaStatusCell.getCellType() == CellType.STRING) ? metaStatusCell.getStringCellValue() : "";

                        if (!metaStatus.equals("Meta product is missing") && !metaStatus.equals("Meta product has errors")) {
                            Row newRow = newSheet.createRow(rowIndex++);

                            for (Map.Entry<String, Integer> entry : columnIndexMap.entrySet()) {
                                String columnName = entry.getKey();
                                Integer destColumnIndex = entry.getValue();
                                Integer sourceColumnIndex = sourceColumnIndexMap.get(columnName);

                                if (sourceColumnIndex != null) {
                                    Cell sourceCell = dataRow.getCell(sourceColumnIndex);
                                    Cell newCell = newRow.createCell(destColumnIndex);

                                    if (sourceCell != null) {
                                        switch (sourceCell.getCellType()) {
                                            case STRING:
                                                newCell.setCellValue(sourceCell.getStringCellValue());
                                                if (columnName.equals("Variant SKU")) {
                                                    variantSKUs.add("('" + sourceCell.getStringCellValue() + "')");
                                                }
                                                break;
                                            case NUMERIC:
                                                newCell.setCellValue(sourceCell.getNumericCellValue());
                                                break;
                                            case BOOLEAN:
                                                newCell.setCellValue(sourceCell.getBooleanCellValue());
                                                break;
                                            case FORMULA:
                                                FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
                                                CellValue cellValue = evaluator.evaluate(sourceCell);
                                                switch (cellValue.getCellType()) {
                                                    case STRING:
                                                        newCell.setCellValue(cellValue.getStringValue());
                                                        if (columnName.equals("Variant SKU")) {
                                                            variantSKUs.add("('" + cellValue.getStringValue() + "')");
                                                        }
                                                        break;
                                                    case NUMERIC:
                                                        newCell.setCellValue(cellValue.getNumberValue());
                                                        break;
                                                    case BOOLEAN:
                                                        newCell.setCellValue(cellValue.getBooleanValue());
                                                        break;
                                                    default:
                                                        newCell.setCellValue("");
                                                        break;
                                                }
                                                break;
                                            default:
                                                newCell.setCellValue("");
                                                break;
                                        }
                                    } else {
                                        newCell.setCellValue("");
                                    }
                                }
                            }
                        }
                    }
                }
            }

            String query = "SELECT \n    p.productcode\nFROM \n    (VALUES " + String.join(", ", variantSKUs) + ") AS p(productcode)\nLEFT JOIN \n    productitem pi ON p.productcode = pi.productcode\nWHERE \n    pi.productcode IS NULL;";

            FileChooser fileChooser = new FileChooser();
            fileChooser.setTitle("Save Corrected Output");
            fileChooser.setInitialFileName("CorrectedOutput.xlsx");
            File outputFile = fileChooser.showSaveDialog(null);


            FileChooser sqlFileChooser = new FileChooser();
            sqlFileChooser.setTitle("Save SQL Query");
            sqlFileChooser.setInitialFileName("query.sql");
            File sqlFile = sqlFileChooser.showSaveDialog(null);



            if (outputFile != null) {
                try (FileOutputStream outputStream = new FileOutputStream(outputFile)) {
                    newWorkbook.write(outputStream);
                    displayInfo("Corrected output saved successfully.");
                }
                processedExcelFiles.add(outputFile);
            } else {
                displayInfo("File save canceled.");
            }

            if (sqlFile != null) {
                try (FileWriter writer = new FileWriter(sqlFile)) {
                    writer.write(query);
                    displayInfo("SQL query saved successfully.");
                } catch (IOException e) {
                    displayError("Error saving SQL file: " + e.getMessage());
                }
            }


        } catch (IOException e) {
            displayError("Error processing file: " + e.getMessage());
        } catch (IllegalArgumentException e) {
            displayError("Error reading Excel file: " + e.getMessage());
        }
    }

    


    // *NEW*: Method to determine the attempt count.
    private int getAttemptCount(String baseName) {
        int count = 1;
        for (File file : processedExcelFiles) {
            if (file.getName().startsWith(baseName + "_attempt_")) {
                // Extract the attempt number and find the highest.
                String name = file.getName();
                String attemptStr = name.substring((baseName + "_attempt_").length(), name.lastIndexOf("."));
                try {
                    int attempt = Integer.parseInt(attemptStr);
                    count = Math.max(count, attempt + 1); // Next attempt number
                } catch (NumberFormatException e) {
                    // Ignore files with incorrectly formatted attempt numbers.
                    System.err.println("Invalid attempt number in filename: " + file.getName());
                }
            }
        }
        return count;
    }



    private void viewExcelFile() {
        if (processedExcelFiles.isEmpty()) {
            displayError("No Excel files processed yet.");
            return;
        }

        // Create a dialog for the user to select which output file to view
        ChoiceDialog<File> choiceDialog = new ChoiceDialog<>(processedExcelFiles.get(0), processedExcelFiles);
        choiceDialog.setTitle("Select Output File");
        choiceDialog.setHeaderText("Choose an output file to view:");
        choiceDialog.setContentText("Output Files:");

        // Apply custom dialog style
        applyDialogStyle(choiceDialog);

        // Show the dialog and wait for the user to select a file
        Optional<File> selectedFile = choiceDialog.showAndWait();
        selectedFile.ifPresent(this::openExcelFile);
    }


    private void openExcelFile(File file) {
        if (Desktop.isDesktopSupported()) {
            Desktop desktop = Desktop.getDesktop();

            try {
                if (desktop.isSupported(Desktop.Action.OPEN)) {
                    desktop.open(file);
                } else {
                    this.displayError("Opening files is not supported on this system.");
                }
            } catch (IOException e) {
                this.displayError("Error opening Excel file: " + e.getMessage());
            }
        } else {
            this.displayError("Desktop is not supported on this platform.");
        }

    }


    private void saveExcelFile() {
        if (processedExcelFiles.isEmpty()) {
            displayError("No Excel files have been processed yet.");
            return;
        }

        // Create a dropdown list of processed files
        ChoiceDialog<File> choiceDialog = new ChoiceDialog<>(processedExcelFiles.get(0), processedExcelFiles);
        choiceDialog.setTitle("Select Output File");
        choiceDialog.setHeaderText("Choose an output file to save:");
        choiceDialog.setContentText("Output Files:");

        // Apply style to the dialog
        applyDialogStyle(choiceDialog);

        // Show the dialog and wait for the user to select a file
        Optional<File> selectedFile = choiceDialog.showAndWait();

        // If the user made a selection, proceed to save the file
        selectedFile.ifPresent(fileToSave -> {
            // FileChooser to let the user choose the save location
            FileChooser fileChooser = new FileChooser();
            fileChooser.setTitle("Save Excel File");
            fileChooser.setInitialFileName(fileToSave.getName()); // Suggest the same name

            // Set extension filter to .xlsx files
            fileChooser.getExtensionFilters().add(new FileChooser.ExtensionFilter("Excel Files", "*.xlsx"));

            // Load the last output folder path from preferences
            String lastOutputFolder = loadPreference(LAST_OUTPUT_FOLDER_KEY, "");
            if (!lastOutputFolder.isEmpty()) {
                File initialDir = new File(lastOutputFolder);
                if (initialDir.exists()) {
                    fileChooser.setInitialDirectory(initialDir);
                } else {
                    // If the stored directory doesn't exist, fallback to the default
                    fileChooser.setInitialDirectory(new File(getDefaultDirectory()));
                    displayError("Stored output directory does not exist: " + lastOutputFolder + ".  Using default.");
                }
            } else {
                // If no preference is stored, use the default directory
                fileChooser.setInitialDirectory(new File(getDefaultDirectory()));
            }

            // Show save dialog and get the file chosen by the user
            File savedFile = fileChooser.showSaveDialog(new Stage());

            // If the user selected a file (i.e., didn't cancel)
            if (savedFile != null) {
                try {
                    // Get the parent directory of the save location
                    File parentDir = savedFile.getParentFile();

                    // Check if the directory exists
                    if (!parentDir.exists()) {
                        displayError("Error: Directory does not exist: " + parentDir.getAbsolutePath());
                        return;
                    }

                    // Check if the directory is writable
                    if (!parentDir.canWrite()) {
                        displayError("Error: No write permission to directory: " + parentDir.getAbsolutePath());
                        return;
                    }

                    // Copy the content of the selected processed file to the saved file
                    Files.copy(fileToSave.toPath(), savedFile.toPath(), StandardCopyOption.REPLACE_EXISTING);
                    displayInfo("Excel file saved to: " + savedFile.getAbsolutePath());
                    savedExcelFiles.add(fileToSave); // Mark the file as saved

                    // Save the last output folder path to preferences
                    savePreference(LAST_OUTPUT_FOLDER_KEY, savedFile.getParent());

                } catch (IOException e) {
                    displayError("Error saving Excel file: " + e.getMessage() + "\nStack Trace:\n" + getStackTraceString(e));
                }
            } else {
                displayInfo("Save operation cancelled by user.");
            }
        });
    }

    private void displayError(String message) {
        Platform.runLater(() -> {
            errorTextArea.appendText("Error: " + message + "\n");
        });
    }

    private void displayInfo(String message) {
        Platform.runLater(() -> {
            errorTextArea.appendText("Info: " + message + "\n");
        });
    }

    private String getStackTraceString(Exception e) {
        StringWriter sw = new StringWriter();
        PrintWriter pw = new PrintWriter(sw);
        e.printStackTrace(pw);
        return sw.toString();
    }


    // Inner class to encapsulate CSV processing logic
    public static class CsvProcessor {

        private static final String[] REQUIRED_HEADERS = {"Handle", "Title", "Product Category", "Option1 Name", "Option1 Value", "Option2 Name", "Option2 Value", "Variant SKU"};

        private static final Set<String> VALID_OPTION_TYPES = new HashSet<>(Arrays.asList("color", "colour", "size", "category", "group", "title"));

        public boolean processCsv(String inputFilePath, String outputFilePath, TextArea errorTextArea) throws IOException {
            try {
                if (!isFileWritable(outputFilePath)) {
                    Platform.runLater(() -> errorTextArea.appendText("Error: The output file '" + outputFilePath + "' is open or locked by another process. Please close it and try again.\n"));
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

                try (BOMInputStream bomInputStream = new BOMInputStream(Files.newInputStream(Paths.get(inputFilePath)));
                     CSVParser parser = new CSVParser(new InputStreamReader(bomInputStream, StandardCharsets.UTF_8),
                             CSVFormat.DEFAULT.withHeader())) {

                // Validate required headers
                    Set<String> headersInFile = new HashSet<>(parser.getHeaderMap().keySet());
                    List<String> missingHeaders = new ArrayList<>();
                    Set<String> headersInFileNormalized = headersInFile.stream()
                            .map(header -> header.trim().toLowerCase()) // Trim spaces & normalize case
                            .collect(Collectors.toSet());

                    for (String requiredHeader : REQUIRED_HEADERS) {
                        if (!headersInFileNormalized.contains(requiredHeader.toLowerCase())) {
                            missingHeaders.add(requiredHeader);
                        }
                    }

                    if (!missingHeaders.isEmpty()) {
                        String errorMessage = "Warning: The following required headers are missing from your CSV file: " + missingHeaders +
                                ". Please update your CSV file headers to include: " + Arrays.toString(REQUIRED_HEADERS);
                        Platform.runLater(() -> errorTextArea.appendText(errorMessage + "\n"));
                        return false; // Returning false to indicate header validation failure
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


        // Helper method to find the header row dynamically
        private int findHeaderRow(Sheet sheet, String[] expectedHeaders) {
            for (Row row : sheet) {
                boolean isHeaderRow = true;
                for (int i = 0; i < expectedHeaders.length; i++) {
                    Cell cell = row.getCell(i);
                    if (cell == null || !cell.getStringCellValue().equals(expectedHeaders[i])) {
                        isHeaderRow = false;
                        break;
                    }
                }
                if (isHeaderRow) {
                    return row.getRowNum(); // Return the index of the header row
                }
            }
            return -1; // Header row not found
        }


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




        public boolean hasErrors(File excelFile) throws IOException {
            try (FileInputStream fileInputStream = new FileInputStream(excelFile);
                 Workbook workbook = new XSSFWorkbook(fileInputStream)) {

                // Define the expected headers for error sheets
                String[] expectedErrorHeaders = {"Error Log", "Handle", "Title", "Product Category", "Option 1 Name", "Option 1 Value", "Option 2 Name", "Option 2 Value", "Variant SKU", "Meta Status"};

                // Check if any error sheet has entries
                String[] errorSheetNames = {"Invalid - Duplicate SKUs", "Invalid Options", "Other Errors"};
                for (String sheetName : errorSheetNames) {
                    Sheet sheet = workbook.getSheet(sheetName);
                    if (sheet != null) {
                        // Find the header row dynamically
                        int headerRowIndex = findHeaderRow(sheet, expectedErrorHeaders);
                        if (headerRowIndex != -1) {
                            // Check if there are any rows after the header row
                            if (sheet.getPhysicalNumberOfRows() > headerRowIndex + 1) {
                                return true; // Errors found
                            }
                        }
                    }
                }

                // Define the expected headers for the success sheet
                String[] expectedSuccessHeaders = {"Handle", "Title", "Product Category", "Option1 Name", "Option1 Value", "Option2 Name", "Option2 Value", "Variant SKU", "Meta Status"};

                // Check if the "Success" sheet has any records with meta status indicating issues
                Sheet successSheet = workbook.getSheet("Success");
                if (successSheet != null) {
                    // Find the header row dynamically
                    int headerRowIndex = findHeaderRow(successSheet, expectedSuccessHeaders);
                    if (headerRowIndex != -1) {
                        // Iterate through rows after the header row
                        for (int i = headerRowIndex + 1; i <= successSheet.getLastRowNum(); i++) {
                            Row row = successSheet.getRow(i);
                            if (row != null) {
                                Cell metaStatusCell = row.getCell(8); // Assuming meta status is in the 9th column (index 8)
                                if (metaStatusCell != null && !metaStatusCell.getStringCellValue().isEmpty()) {
                                    return true; // Meta issues found
                                }
                            }
                        }
                    }
                }

                return false; // No errors found
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


