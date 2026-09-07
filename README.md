# CSV → Excel Generator

Desktop JavaFX app that validates Shopify product CSVs and writes structured Excel reports (errors + success sheets). Also supports variation/product upload templates and Apple Numbers → CSV conversion via Aspose Cells.

## What’s improved in 1.1.0

- **Null / empty Title safety** — missing or null `Title` (and related cells) no longer crash the app; they are validated and reported as before.
- **Clearer workflow** — Select → Process → Review & Save.
- **Modern UI** — charcoal / teal desktop theme.
- **Performance** — single-pass Excel write (errors + success in one workbook), background processing with progress, no artificial delays.
- **Stronger error handling** — bad handles/rows don’t abort the whole file; clearer activity log messages.
- **macOS & Linux packaging scripts** via `jpackage`.

## Requirements

- JDK **17+** (includes `jpackage`)
- Maven 3.8+
- Network access for Aspose Maven repo (Numbers conversion)

## Run (dev)

```bash
mvn javafx:run
```

Or package a fat jar:

```bash
mvn -DskipTests package
java -jar target/csv-to-excel-generator-1.1.0.jar
```

> Fat jars that embed JavaFX may need platform-specific JavaFX natives. Prefer `mvn javafx:run` for local use, or the `jpackage` scripts below for distributable apps.

## Package

### macOS

```bash
./scripts/package-macos.sh
# → dist/macos/CSV to Excel Generator.app
```

### Linux

```bash
./scripts/package-linux.sh
# → dist/linux/csv-to-excel-generator/  (app-image)
# → dist/linux/*.deb when dpkg tools are available
```

Cross-building: run the macOS script on a Mac and the Linux script on Linux (jpackage cannot cross-OS).

## Workflow

1. **Select** one or more Shopify CSV files (or use template / Numbers tools).
2. **Process** to validate handles, titles, options, and SKUs.
3. **Review & Save** the generated Excel report; optionally generate corrected output.

Empty titles on suspected meta products are flagged in **Other Errors** instead of throwing.

## Note on sample CSVs

Large sample inputs in the repo root are for local testing; prefer keeping them out of commits when possible.
