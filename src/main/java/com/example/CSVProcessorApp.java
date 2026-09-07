
package com.example;

import com.aspose.cells.SaveFormat;
import javafx.animation.Animation;
import javafx.animation.FadeTransition;
import javafx.animation.KeyFrame;
import javafx.animation.KeyValue;
import javafx.animation.Timeline;
import javafx.application.Application;
import javafx.scene.effect.BlendMode;
import javafx.scene.effect.DropShadow;
import javafx.scene.paint.CycleMethod;
import javafx.scene.paint.LinearGradient;
import javafx.scene.paint.Stop;
import javafx.scene.shape.Rectangle;
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
import javafx.stage.Window;
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
    private Stage primaryStage;


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
        this.primaryStage = primaryStage;
        primaryStage.setTitle("CSV → Excel Generator");

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

        uploadFileTypeComboBox.setPromptText("Choose Variation or Product…");
        uploadFileTypeComboBox.setPrefWidth(240);


        processTemplate1 = new Button("process Template");



        saveTemplate1 = new Button("save Template");

        clearSelectedFilesButton.setOnAction(e -> {
            System.out.println("Clear all Button Clicked");
            clearSelectedFiles();
        });
        Button correctedOutputButton = new Button("Corrected Output");
        correctedOutputButton.setPrefWidth(180);
        correctedOutputButton.setOnAction(e -> {
            System.out.println("Corrected Output Button Clicked");
            chooseProcessedFileAndGenerateCorrectedOutput();
        });

        Button generateSqlButton = new Button("Generate SQL");
        generateSqlButton.setPrefWidth(160);
        generateSqlButton.setOnAction(e -> {
            System.out.println("Generate SQL Button Clicked");
            chooseProcessedFileAndGenerateSql();
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
                } else if ("Product Upload File".equals(selected)) {
                    processProductUpload();
                    if (!processedExcelFiles.isEmpty() && processedExcelFiles.get(0).exists()) {
                        processTemplate1.setText("view product upload");
                    }
                } else {
                    displayError("Choose a Variation or Product template first.");
                }
            } catch (IOException ex) {
                displayError("Error processing template: " + ex.getMessage());
            }
        });




        selectCsvButton.getStyleClass().add("primary-button");
        processButton.getStyleClass().add("primary-button");
        viewExcelButton.getStyleClass().add("ghost-button");
        saveExcelButton.getStyleClass().add("ghost-button");
        correctedOutputButton.getStyleClass().add("ghost-button");
        generateSqlButton.getStyleClass().add("ghost-button");
        clearSelectedFilesButton.getStyleClass().add("danger-button");
        convertNumbersToCsvButton.getStyleClass().add("ghost-button");

        Label brandTitle = new Label("CSV → Excel Generator");
        brandTitle.getStyleClass().add("brand-title");
        Label brandSubtitle = new Label("Shopify product CSV validation and Excel reporting");
        brandSubtitle.getStyleClass().add("brand-subtitle");
        VBox brandBox = new VBox(4, brandTitle, brandSubtitle);
        brandBox.getStyleClass().add("brand-box");

        Label step1Title = new Label("1 · Select");
        step1Title.getStyleClass().add("step-title");
        Label step2Title = new Label("2 · Process");
        step2Title.getStyleClass().add("step-title");
        Label step3Title = new Label("3 · Review & Save");
        step3Title.getStyleClass().add("step-title");

        HBox processContainer = new HBox(10);
        processContainer.setAlignment(Pos.CENTER_LEFT);
        processContainer.setPadding(new Insets(8, 0, 8, 0));
        processContainer.getChildren().addAll(processButton);

        FlowPane reviewContainer = new FlowPane(10, 10);
        reviewContainer.setAlignment(Pos.CENTER_LEFT);
        reviewContainer.setPadding(new Insets(8, 0, 8, 0));
        reviewContainer.getChildren().addAll(viewExcelButton, saveExcelButton, correctedOutputButton, generateSqlButton);

        Label pathATitle = new Label("A · Validate a CSV  (direct select)");
        pathATitle.getStyleClass().add("path-title");
        Label pathAHint = new Label("Shopify product export → check errors → Excel / SQL");
        pathAHint.getStyleClass().add("status-label");
        HBox pathAActions = new HBox(10, selectCsvButton);
        VBox pathA = new VBox(6, pathATitle, pathAHint, pathAActions);
        pathA.getStyleClass().add("path-card");

        Label pathBTitle = new Label("B · Build an upload file  (template)");
        pathBTitle.getStyleClass().add("path-title");
        Label pathBHint = new Label("Same CSV, but Variation or Product upload format");
        pathBHint.getStyleClass().add("status-label");
        FlowPane pathBActions = new FlowPane(10, 8);
        pathBActions.getChildren().addAll(uploadFileTypeComboBox, processTemplate1, saveTemplate1);
        VBox pathB = new VBox(6, pathBTitle, pathBHint, pathBActions);
        pathB.getStyleClass().add("path-card");

        Label pathCTitle = new Label("C · Apple Numbers  (.numbers → .csv)");
        pathCTitle.getStyleClass().add("path-title");
        Label pathCHint = new Label("Convert first, then use the CSV in A or B");
        pathCHint.getStyleClass().add("status-label");
        HBox pathCActions = new HBox(10, convertNumbersToCsvButton);
        VBox pathC = new VBox(6, pathCTitle, pathCHint, pathCActions);
        pathC.getStyleClass().add("path-card");

        Label guidanceBody = new Label(
                "• Input your CSV file to check validations, press Process, then you can Save, View,\n" +
                "  generate a Corrected XLSX, or Generate SQL.\n" +
                "• To create a Variation or Product upload file from the CSV: select the template,\n" +
                "  Process Template, then Save Template.\n" +
                "• Numbers files: convert .numbers → .csv first, then use the CSV above."
        );
        guidanceBody.setWrapText(true);
        guidanceBody.getStyleClass().add("guidance-body");
        guidanceBody.setManaged(false);
        guidanceBody.setVisible(false);

        Label hintArrow = new Label("▼");
        hintArrow.getStyleClass().add("hint-arrow");
        Label hintText = new Label("What to select");
        hintText.getStyleClass().add("hint-text");
        HBox hintContent = new HBox(8, hintArrow, hintText);
        hintContent.setAlignment(Pos.CENTER_LEFT);
        StackPane hintStack = wrapWithShimmer(hintContent);
        hintStack.getStyleClass().add("select-hint");
        hintStack.setCursor(javafx.scene.Cursor.HAND);
        hintStack.setOnMouseClicked(e -> {
            boolean show = !guidanceBody.isVisible();
            guidanceBody.setVisible(show);
            guidanceBody.setManaged(show);
            hintArrow.setText(show ? "▲" : "▼");
        });

        HBox step1Header = new HBox();
        step1Header.setAlignment(Pos.CENTER_LEFT);
        Region headerSpacer = new Region();
        HBox.setHgrow(headerSpacer, Priority.ALWAYS);
        step1Header.getChildren().addAll(step1Title, headerSpacer, clearSelectedFilesButton);

        VBox step1Panel = new VBox(10, step1Header, hintStack, guidanceBody, pathA, pathB, pathC);
        step1Panel.getStyleClass().add("panel");
        Label step2Hint = new Label("For path A: validate the CSV selected above. Progress appears below.");
        step2Hint.getStyleClass().add("status-label");
        VBox step2Panel = new VBox(8, step2Title, step2Hint, processContainer);
        step2Panel.getStyleClass().add("panel");
        VBox step3Panel = new VBox(8, step3Title, reviewContainer);
        step3Panel.getStyleClass().add("panel");

        variationSelectedFileLabel= new Label("No Variation upload file selected");
        variationSelectedFileLabel.getStyleClass().add("status-label");
        selectedFileLabel = new Label("No CSV file selected");
        selectedFileLabel.getStyleClass().add("status-label");
        errorTextArea = new TextArea();
        errorTextArea.setEditable(false);
        errorTextArea.setPrefHeight(180);
        errorTextArea.getStyleClass().add("console");
        progressBar = new ProgressBar(0);
        progressBar.setMaxWidth(Double.MAX_VALUE);
        progressBar.setVisible(false);
        progressLabel = new Label("");
        progressLabel.getStyleClass().add("status-label");

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


        selectCsvButton.setPrefWidth(180);
        processButton.setPrefWidth(180);
        viewExcelButton.setPrefWidth(150);
        saveExcelButton.setPrefWidth(150);
        correctedOutputButton.setPrefWidth(170);
        generateSqlButton.setPrefWidth(150);




        Label messagesTitle = new Label("Activity log");
        messagesTitle.getStyleClass().add("section-label");

        VBox layout = new VBox(14);
        layout.setPadding(new Insets(20));
        layout.getStyleClass().add("root-layout");
        layout.getChildren().addAll(
                brandBox,
                step1Panel,
                step2Panel,
                step3Panel,
                selectedFileLabel,
                variationSelectedFileLabel,
                progressLabel,
                progressBar,
                messagesTitle,
                errorTextArea,
                toggleInstructionsButton,
                instructionsLabel
        );

        javafx.scene.control.ScrollPane scroller = new javafx.scene.control.ScrollPane(layout);
        scroller.setFitToWidth(true);
        scroller.setHbarPolicy(javafx.scene.control.ScrollPane.ScrollBarPolicy.NEVER);
        scroller.getStyleClass().add("root-layout");

        Scene scene = new Scene(scroller, 1020, 820);

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
            String recentInputFolderForNumbers = loadPreference(RECENT_INPUT_FOLDER_KEY, "");
            if (!recentInputFolderForNumbers.isEmpty()) {
                File recentDir = new File(recentInputFolderForNumbers);
                if (recentDir.isDirectory()) {
                    fileChooser.setInitialDirectory(recentDir);
                }
            }
            File numbersFile = fileChooser.showOpenDialog(ownerWindow());
            if (numbersFile == null) {
                displayError("No .numbers file selected.");
                return;
            }

            FileChooser saveChooser = new FileChooser();
            saveChooser.setTitle("Save Converted CSV");
            saveChooser.getExtensionFilters().add(new FileChooser.ExtensionFilter("CSV Files (*.csv)", "*.csv"));
            String base = numbersFile.getName().replaceFirst("\\.numbers$", "");
            saveChooser.setInitialFileName(base + ".csv");
            if (numbersFile.getParentFile() != null) {
                saveChooser.setInitialDirectory(numbersFile.getParentFile());
            }
            File csvFile = saveChooser.showSaveDialog(ownerWindow());
            if (csvFile == null) {
                displayError("No save location selected.");
                return;
            }

            try {
                convertNumbersToCsv(numbersFile, csvFile);
                displayInfo("Conversion successful! Saved to: " + csvFile.getAbsolutePath());
                savePreference(RECENT_INPUT_FOLDER_KEY, csvFile.getParent());
            } catch (Exception ex) {
                ex.printStackTrace();
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
        Exception asposeError = null;
        try {
            com.aspose.cells.Workbook workbook = new com.aspose.cells.Workbook(numbersFile.getAbsolutePath());
            workbook.save(csvFile.getAbsolutePath(), com.aspose.cells.SaveFormat.CSV);
            if (csvFile.exists() && csvFile.length() > 0) {
                return;
            }
        } catch (Exception e) {
            asposeError = e;
        }

        if (System.getProperty("os.name", "").toLowerCase().contains("mac")) {
            convertNumbersViaNumbersApp(numbersFile, csvFile);
            return;
        }

        if (asposeError != null) {
            throw new Exception(
                    "Could not convert this .numbers file. On a Mac, install Apple Numbers and try again. " +
                            "Otherwise export CSV from Numbers, or use https://cloudconvert.com/numbers-to-csv. " +
                            "Detail: " + asposeError.getMessage(),
                    asposeError
            );
        }
        throw new Exception("Conversion produced an empty file.");
    }

    private void convertNumbersViaNumbersApp(File numbersFile, File csvFile) throws Exception {
        File numbersApp = new File("/Applications/Numbers.app");
        if (!numbersApp.exists()) {
            throw new Exception("Apple Numbers is not installed at /Applications/Numbers.app, so this .numbers file cannot be converted automatically.");
        }

        String inPath = numbersFile.getAbsolutePath().replace("\\", "\\\\").replace("\"", "\\\"");
        String outPath = csvFile.getAbsolutePath().replace("\\", "\\\\").replace("\"", "\\\"");
        String script =
                "tell application \"Numbers\"\n" +
                "  open POSIX file \"" + inPath + "\"\n" +
                "  delay 1.5\n" +
                "  export front document to POSIX file \"" + outPath + "\" as CSV\n" +
                "  close front document saving no\n" +
                "end tell\n";

        Process process = new ProcessBuilder("osascript", "-").redirectErrorStream(true).start();
        try (OutputStreamWriter writer = new OutputStreamWriter(process.getOutputStream(), StandardCharsets.UTF_8)) {
            writer.write(script);
        }
        int code = process.waitFor();
        String output;
        try (BufferedReader reader = new BufferedReader(new InputStreamReader(process.getInputStream(), StandardCharsets.UTF_8))) {
            StringBuilder sb = new StringBuilder();
            String line;
            while ((line = reader.readLine()) != null) {
                sb.append(line).append('\n');
            }
            output = sb.toString().trim();
        }

        File produced = csvFile;
        if (!produced.exists() || produced.length() == 0) {
            File asFolder = csvFile;
            if (asFolder.isDirectory()) {
                File[] csvs = asFolder.listFiles((dir, name) -> name.toLowerCase().endsWith(".csv"));
                if (csvs != null && csvs.length > 0) {
                    Files.copy(csvs[0].toPath(), csvFile.toPath(), StandardCopyOption.REPLACE_EXISTING);
                    produced = csvFile;
                }
            }
        }

        if (code != 0 || !csvFile.exists() || csvFile.length() == 0) {
            throw new Exception("Numbers export failed" + (output.isEmpty() ? "." : ": " + output));
        }
    }

    private StackPane wrapWithShimmer(Node content) {
        content.setMouseTransparent(false);
        StackPane stack = new StackPane(content);
        stack.setAlignment(Pos.CENTER_LEFT);
        stack.setMaxWidth(Region.USE_PREF_SIZE);

        Rectangle shimmer = new Rectangle(90, 18);
        shimmer.setMouseTransparent(true);
        shimmer.setBlendMode(BlendMode.ADD);
        shimmer.setFill(new LinearGradient(
                0, 0, 1, 0, true, CycleMethod.NO_CYCLE,
                new Stop(0, Color.TRANSPARENT),
                new Stop(0.35, Color.web("#5eead4", 0.05)),
                new Stop(0.5, Color.web("#99f6e4", 0.85)),
                new Stop(0.65, Color.web("#5eead4", 0.05)),
                new Stop(1, Color.TRANSPARENT)
        ));
        stack.getChildren().add(shimmer);

        Rectangle clip = new Rectangle();
        clip.widthProperty().bind(stack.widthProperty());
        clip.heightProperty().bind(stack.heightProperty());
        clip.setArcWidth(8);
        clip.setArcHeight(8);
        stack.setClip(clip);

        DropShadow glow = new DropShadow();
        glow.setColor(Color.web("#5eead4", 0.55));
        glow.setRadius(14);
        content.setEffect(glow);

        shimmer.heightProperty().bind(stack.heightProperty());
        Timeline sweep = new Timeline();
        sweep.setCycleCount(Animation.INDEFINITE);
        sweep.setAutoReverse(true);
        Runnable rebuild = () -> {
            double width = Math.max(stack.getWidth(), 180);
            sweep.stop();
            sweep.getKeyFrames().setAll(
                    new KeyFrame(Duration.ZERO, new KeyValue(shimmer.translateXProperty(), -90)),
                    new KeyFrame(Duration.seconds(1.6), new KeyValue(shimmer.translateXProperty(), width))
            );
            sweep.playFromStart();
        };
        stack.widthProperty().addListener((obs, oldW, newW) -> rebuild.run());
        Platform.runLater(rebuild);
        return stack;
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




    private Window ownerWindow() {
        return primaryStage != null ? primaryStage : null;
    }

    private void styleOwnedDialog(Dialog<?> dialog) {
        applyDialogStyle(dialog);
        if (primaryStage != null) {
            dialog.initOwner(primaryStage);
        }
    }

    private File resolveOutputFile(File csvFile, String baseName, int attemptCount) {
        File dir = csvFile.getParentFile();
        if (dir == null || !dir.canWrite()) {
            dir = new File(System.getProperty("user.home"), "Documents/CSVExcelGenerator");
            if (!dir.exists() && !dir.mkdirs()) {
                dir = new File(System.getProperty("java.io.tmpdir"), "CSVExcelGenerator");
                dir.mkdirs();
            }
        }
        return new File(dir, baseName + "_attempt_" + attemptCount + ".xlsx");
    }

    private void selectCsvFile() {
        FileChooser fileChooser = new FileChooser();
        fileChooser.setTitle("Select CSV Files");
        fileChooser.getExtensionFilters().add(new FileChooser.ExtensionFilter("CSV Files", "*.csv"));

        // Set the initial directory from preferences
        String recentInputFolder = loadPreference(RECENT_INPUT_FOLDER_KEY, "");
        if (!recentInputFolder.isEmpty()) {
            File recentDir = new File(recentInputFolder);
            if (recentDir.isDirectory()) {
                fileChooser.setInitialDirectory(recentDir);
            }
        }

        List<File> selectedFiles = fileChooser.showOpenMultipleDialog(ownerWindow());

        if (selectedFiles != null && !selectedFiles.isEmpty()) {
            selectedCsvFiles.clear(); // Clear previous selections
            selectedCsvFiles.addAll(selectedFiles);
            String names = selectedFiles.stream().map(File::getName).collect(Collectors.joining(", "));
            selectedFileLabel.setText(selectedFiles.size() + " CSV selected: " + names);

            // Save the recent input folder path
            savePreference(RECENT_INPUT_FOLDER_KEY, selectedFiles.get(0).getParent());
        } else {
            selectedCsvFiles.clear();
            selectedFileLabel.setText("No CSV file selected");
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

        progressBar.setVisible(true);
        progressBar.setProgress(0);
        progressLabel.setText("Starting…");

        Task<Void> processingTask = new Task<Void>() {
            @Override
            protected Void call() throws Exception {
                int totalFiles = selectedCsvFiles.size();
                for (int i = 0; i < totalFiles; i++) {
                    File csvFile = selectedCsvFiles.get(i);
                    String inputFilePath = csvFile.getAbsolutePath();
                    String baseName = csvFile.getName().replaceFirst("[.][^.]+$", ""); // Filename without extension

                    int attemptCount = getAttemptCount(baseName);
                    File outputFile = resolveOutputFile(csvFile, baseName, attemptCount);
                    String outputFilePath = outputFile.getAbsolutePath();

                    final int fileIndex = i + 1;
                    updateMessage("Processing " + csvFile.getName() + " (" + fileIndex + "/" + totalFiles + ")");
                    updateProgress(fileIndex - 1, totalFiles);
                    try {
                        boolean success = csvProcessor.processCsv(inputFilePath, outputFilePath, errorTextArea);
                        if (success) {
                            processedExcelFiles.add(outputFile.getAbsoluteFile());
                            Platform.runLater(() -> displayInfo("Output ready: " + outputFile.getAbsolutePath()));

                            boolean hasErrors = csvProcessor.hasErrors(outputFile);
                            if (hasErrors) {
                                Platform.runLater(() -> displayError("Validation issues found — review the Excel report: " + csvFile.getName()));
                            } else {
                                Platform.runLater(() -> displayInfo("No validation errors: " + csvFile.getName()));
                            }
                        } else {
                            Platform.runLater(() -> displayError("Could not process: " + csvFile.getName()));
                        }
                    } catch (Exception e) {
                        Platform.runLater(() -> displayError("Error processing " + csvFile.getName() + ": " +
                                (e.getMessage() != null ? e.getMessage() : e.toString())));
                    }
                    updateProgress(fileIndex, totalFiles);
                            }

                            return null;
                        }

            @Override
            protected void succeeded() {
                super.succeeded();
                Platform.runLater(() -> {
                    displayInfo("All files processed.");
                    progressBar.progressProperty().unbind();
                    progressBar.setProgress(1);
                    progressBar.setVisible(false);
                    progressLabel.setText("Done");
                });
            }

            @Override
            protected void failed() {
                super.failed();
                Throwable error = getException();
                Platform.runLater(() -> {
                    displayError("File processing failed: " + (error != null ? error.getMessage() : "Unknown error"));
                    progressBar.progressProperty().unbind();
                    progressBar.setVisible(false);
                    progressLabel.setText("Failed");
                });
            }
        };

        progressBar.progressProperty().bind(processingTask.progressProperty());
        processingTask.messageProperty().addListener((obs, o, n) -> progressLabel.setText(n));
        Thread worker = new Thread(processingTask, "csv-processor");
        worker.setDaemon(true);
        worker.start();
    }


    private void chooseProcessedFileAndGenerateCorrectedOutput() {
        chooseProcessedFile("Select the processed Excel file to generate corrected output from:")
                .ifPresent(this::generateCorrectedOutput);
    }

    private void chooseProcessedFileAndGenerateSql() {
        chooseProcessedFile("Select the processed Excel file to generate SQL from:")
                .ifPresent(this::generateSqlQuery);
    }

    private Optional<File> chooseProcessedFile(String headerText) {
        processedExcelFiles.removeIf(f -> f == null || !f.exists());
        if (processedExcelFiles.isEmpty()) {
            displayError("No processed Excel files available. Process a CSV first.");
            return Optional.empty();
        }

        ChoiceDialog<File> dialog = new ChoiceDialog<>(processedExcelFiles.get(0), processedExcelFiles);
        dialog.setTitle("Choose Processed File");
        dialog.setHeaderText(headerText);
        dialog.setContentText("Choose a file:");
        styleOwnedDialog(dialog);
        return dialog.showAndWait();
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

            FileChooser fileChooser = new FileChooser();
            fileChooser.setTitle("Save Corrected Output");
            fileChooser.setInitialFileName("CorrectedOutput.xlsx");
            fileChooser.getExtensionFilters().add(new FileChooser.ExtensionFilter("Excel Files", "*.xlsx"));
            File outputFile = fileChooser.showSaveDialog(ownerWindow());

            if (outputFile != null) {
                try (FileOutputStream outputStream = new FileOutputStream(outputFile)) {
                    newWorkbook.write(outputStream);
                    displayInfo("Corrected output saved successfully: " + outputFile.getAbsolutePath());
                }
                processedExcelFiles.add(outputFile.getAbsoluteFile());
            } else {
                displayInfo("Corrected output save canceled.");
            }
            newWorkbook.close();

        } catch (IOException e) {
            displayError("Error processing file: " + e.getMessage());
        } catch (IllegalArgumentException e) {
            displayError("Error reading Excel file: " + e.getMessage());
        }
    }

    private void generateSqlQuery(File selectedExcelFile) {
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

            Map<String, Integer> sourceColumnIndexMap = new HashMap<>();
            Row headerRowSource = successSheet.getRow(1);
            if (headerRowSource != null) {
                for (int i = 0; i < headerRowSource.getLastCellNum(); i++) {
                    Cell cell = headerRowSource.getCell(i);
                    if (cell != null) {
                        sourceColumnIndexMap.put(cell.getStringCellValue().trim(), i);
                    }
                }
            }

            if (!sourceColumnIndexMap.containsKey("Variant SKU") || !sourceColumnIndexMap.containsKey("Meta Status")) {
                displayError("Success sheet is missing Variant SKU or Meta Status columns.");
                return;
            }

            List<String> variantSKUs = new ArrayList<>();
            int skuCol = sourceColumnIndexMap.get("Variant SKU");
            int metaCol = sourceColumnIndexMap.get("Meta Status");

            for (int i = 2; i <= successSheet.getLastRowNum(); i++) {
                Row dataRow = successSheet.getRow(i);
                if (dataRow == null) {
                    continue;
                }
                Cell metaStatusCell = dataRow.getCell(metaCol);
                String metaStatus = "";
                if (metaStatusCell != null && metaStatusCell.getCellType() == CellType.STRING) {
                    metaStatus = metaStatusCell.getStringCellValue();
                }
                if (metaStatus.equals("Meta product is missing") || metaStatus.equals("Meta product has errors")) {
                    continue;
                }
                Cell skuCell = dataRow.getCell(skuCol);
                if (skuCell == null) {
                    continue;
                }
                String sku = "";
                if (skuCell.getCellType() == CellType.STRING) {
                    sku = skuCell.getStringCellValue();
                } else if (skuCell.getCellType() == CellType.NUMERIC) {
                    sku = String.valueOf((long) skuCell.getNumericCellValue());
                }
                if (sku != null && !sku.trim().isEmpty()) {
                    variantSKUs.add("('" + sku.trim().replace("'", "''") + "')");
                }
            }

            if (variantSKUs.isEmpty()) {
                displayError("No valid Variant SKUs found to build SQL.");
                return;
            }

            String query = "SELECT \n    p.productcode\nFROM \n    (VALUES " + String.join(", ", variantSKUs)
                    + ") AS p(productcode)\nLEFT JOIN \n    productitem pi ON p.productcode = pi.productcode\nWHERE \n    pi.productcode IS NULL;";

            FileChooser sqlFileChooser = new FileChooser();
            sqlFileChooser.setTitle("Save SQL Query");
            sqlFileChooser.setInitialFileName("query.sql");
            sqlFileChooser.getExtensionFilters().add(new FileChooser.ExtensionFilter("SQL Files", "*.sql"));
            File sqlFile = sqlFileChooser.showSaveDialog(ownerWindow());

            if (sqlFile != null) {
                try (FileWriter writer = new FileWriter(sqlFile)) {
                    writer.write(query);
                    displayInfo("SQL query saved successfully: " + sqlFile.getAbsolutePath());
                }
            } else {
                displayInfo("SQL save canceled.");
            }
        } catch (IOException e) {
            displayError("Error generating SQL: " + e.getMessage());
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
        chooseProcessedFile("Choose an output file to view:")
                .ifPresent(this::openExcelFile);
    }


    private void openExcelFile(File file) {
        if (file == null || !file.exists()) {
            displayError("Output file not found: " + (file == null ? "(null)" : file.getAbsolutePath()));
            return;
        }

        try {
            // Prefer OS open command — more reliable than AWT Desktop inside jpackage apps
            String os = System.getProperty("os.name", "").toLowerCase();
            ProcessBuilder pb;
            if (os.contains("mac")) {
                pb = new ProcessBuilder("open", file.getAbsolutePath());
            } else if (os.contains("win")) {
                pb = new ProcessBuilder("cmd", "/c", "start", "", file.getAbsolutePath());
            } else {
                pb = new ProcessBuilder("xdg-open", file.getAbsolutePath());
            }
            pb.start();
            displayInfo("Opened: " + file.getAbsolutePath());
            return;
        } catch (Exception ignored) {
            // fall through
        }

        try {
            getHostServices().showDocument(file.toURI().toString());
            displayInfo("Opened via HostServices: " + file.getAbsolutePath());
            return;
        } catch (Exception ignored) {
            // fall through
        }

        try {
            if (Desktop.isDesktopSupported() && Desktop.getDesktop().isSupported(Desktop.Action.OPEN)) {
                Desktop.getDesktop().open(file);
                displayInfo("Opened via Desktop: " + file.getAbsolutePath());
            } else {
                displayError("Unable to open file automatically. Path: " + file.getAbsolutePath());
            }
        } catch (IOException e) {
            displayError("Error opening Excel file: " + e.getMessage() + "\nPath: " + file.getAbsolutePath());
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
        styleOwnedDialog(choiceDialog);

        Optional<File> selectedFile = choiceDialog.showAndWait();

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
            File savedFile = fileChooser.showSaveDialog(ownerWindow());

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
                        boolean isImageEntry = getCellValue(record, "Option1 Name").isEmpty() &&
                                getCellValue(record, "Option1 Value").isEmpty() &&
                                getCellValue(record, "Option2 Name").isEmpty() &&
                                getCellValue(record, "Option2 Value").isEmpty() &&
                                getCellValue(record, "Variant SKU").isEmpty();

                        if (!isImageEntry) {
                            recordsToProcess.add(record);
                        } else {
                            imageEntries.add(record);
                        }
                    }

                    for (CSVRecord record : recordsToProcess) {
                        String handle = getCellValue(record, "Handle");
                        handleToRecordsMap.computeIfAbsent(handle, k -> new ArrayList<>()).add(record);
                    }

                    // Process each handle group
                    for (Map.Entry<String, List<CSVRecord>> entry : handleToRecordsMap.entrySet()) {
                        String handle = entry.getKey();
                        List<CSVRecord> records = entry.getValue();
                        try {

                        // Identify "Title/Default Title" Meta Products
                        List<CSVRecord> titleDefaultMetaProducts = records.stream()
                                .filter(r -> getCellValue(r, "Option1 Name").equalsIgnoreCase("Title") &&
                                        getCellValue(r, "Option1 Value").equalsIgnoreCase("Default Title"))
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
                                .filter(r -> !getCellValue(r, "Title").isEmpty())
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
                                currentRecordErrors.add(new ProductError("Missing SKU", record, metaRecord != null ? getCellValue(metaRecord, "Title") : ""));
                                errors.get("Invalid - Duplicate SKUs").add(currentRecordErrors.get(currentRecordErrors.size() - 1));
                            } else if (skuSet.contains(sku)) {
                                currentRecordErrors.add(new ProductError("Duplicate SKU found", record, metaRecord != null ? getCellValue(metaRecord, "Title") : ""));
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

                        } catch (RuntimeException rowEx) {
                            String msg = "Unexpected error while processing handle '" + handle + "': " +
                                    (rowEx.getMessage() != null ? rowEx.getMessage() : rowEx.getClass().getSimpleName());
                            if (records.isEmpty()) {
                                Platform.runLater(() -> errorTextArea.appendText(msg + "\n"));
                            } else {
                                errors.get("Other Errors").add(new ProductError(msg, records.get(0)));
                            }
                        }
                    }
                }

                System.out.println("Skipped image entries: " + imageEntries.size());

                writeResultsToExcel(outputFilePath, errors, successfulRecords);
                System.out.println("Processing completed. Results written to: " + outputFilePath);
            } catch (IOException e) {
                e.printStackTrace();
                String msg = e.getMessage() != null ? e.getMessage() : e.toString();
                Platform.runLater(() -> errorTextArea.appendText("I/O error: " + msg + "\n"));
                return false;
            } catch (RuntimeException e) {
                e.printStackTrace();
                String msg = e.getMessage() != null ? e.getMessage() : e.toString();
                Platform.runLater(() -> errorTextArea.appendText("Processing error: " + msg + "\n"));
                return false;
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


        private static void writeResultsToExcel(String outputFilePath,
                                                 Map<String, List<ProductError>> errors,
                                                 List<SuccessfulRecord> successfulRecords) throws IOException {
            try (Workbook workbook = new XSSFWorkbook()) {
                for (Map.Entry<String, List<ProductError>> entry : errors.entrySet()) {
                    writeErrorsToSheet(workbook, entry.getKey(), entry.getValue());
                }
                writeSuccessfulRecordsToSheet(workbook, successfulRecords);
                try (FileOutputStream outputStream = new FileOutputStream(outputFilePath)) {
                    workbook.write(outputStream);
                }
            } catch (FileNotFoundException e) {
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
                setStringCell(row, 0, error.errorLog);
                setStringCell(row, 1, error.handle);
                setStringCell(row, 2, error.title);
                setStringCell(row, 3, error.productCategory);
                setStringCell(row, 4, error.option1Name);
                setStringCell(row, 5, error.option1Value);
                setStringCell(row, 6, error.option2Name);
                setStringCell(row, 7, error.option2Value);
                setStringCell(row, 8, error.variantSKU);
                setStringCell(row, 9, error.metaStatus);
            }
        }

        private static void writeSuccessfulRecordsToSheet(Workbook workbook, List<SuccessfulRecord> successfulRecords) {
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
                setStringCell(row, 0, getCellValue(record, "Handle"));
                setStringCell(row, 1, getCellValue(record, "Title"));
                setStringCell(row, 2, getCellValue(record, "Product Category"));
                setStringCell(row, 3, getCellValue(record, "Option1 Name"));
                setStringCell(row, 4, getCellValue(record, "Option1 Value"));
                setStringCell(row, 5, getCellValue(record, "Option2 Name"));
                setStringCell(row, 6, getCellValue(record, "Option2 Value"));
                setStringCell(row, 7, getCellValue(record, "Variant SKU"));
                setStringCell(row, 8, successfulRecord.metaStatus);
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
                String value = record.get(headerName);
                return value == null ? "" : value.trim();
            } catch (IllegalArgumentException | IllegalStateException e) {
                return "";
            }
        }

        private static void setStringCell(Row row, int index, String value) {
            row.createCell(index).setCellValue(value == null ? "" : value);
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
                this.errorLog = errorLog == null ? "" : errorLog;
                this.handle = getCellValue(record, "Handle");
                this.title = getCellValue(record, "Title");
                this.productCategory = getCellValue(record, "Product Category");
                this.option1Name = getCellValue(record, "Option1 Name");
                this.option1Value = getCellValue(record, "Option1 Value");
                this.option2Name = getCellValue(record, "Option2 Name");
                this.option2Value = getCellValue(record, "Option2 Value");
                this.variantSKU = getCellValue(record, "Variant SKU");
                this.metaStatus = "";
            }

            public ProductError(String errorLog, CSVRecord record, String metaTitle) {
                this(errorLog, record);
                this.title = metaTitle == null ? "" : metaTitle.trim();
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


