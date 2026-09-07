#!/usr/bin/env bash
set -euo pipefail
ROOT="$(cd "$(dirname "$0")/.." && pwd)"
cd "$ROOT"
mvn -DskipTests package
APP_NAME="CSV to Excel Generator"
JAR=$(ls target/csv-to-excel-generator-*.jar | head -1)
DEST="$ROOT/dist/macos"
rm -rf "$DEST"
mkdir -p "$DEST"
jpackage \
  --type app-image \
  --name "$APP_NAME" \
  --input target \
  --main-jar "$(basename "$JAR")" \
  --main-class com.example.CSVProcessorApp \
  --dest "$DEST" \
  --java-options "-Dfile.encoding=UTF-8" \
  --app-version 1.1.0
echo "macOS app image: $DEST/$APP_NAME.app"
