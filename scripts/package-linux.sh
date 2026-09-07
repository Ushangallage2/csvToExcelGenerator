#!/usr/bin/env bash
set -euo pipefail
ROOT="$(cd "$(dirname "$0")/.." && pwd)"
cd "$ROOT"
mvn -DskipTests package
APP_NAME="csv-to-excel-generator"
JAR=$(ls target/csv-to-excel-generator-*.jar | head -1)
DEST="$ROOT/dist/linux"
rm -rf "$DEST"
mkdir -p "$DEST"
# Prefer app-image (works without fakeroot); also try deb if available
jpackage \
  --type app-image \
  --name "$APP_NAME" \
  --input target \
  --main-jar "$(basename "$JAR")" \
  --main-class com.example.CSVProcessorApp \
  --dest "$DEST" \
  --java-options "-Dfile.encoding=UTF-8" \
  --app-version 1.1.0
if command -v dpkg-deb >/dev/null 2>&1; then
  jpackage \
    --type deb \
    --name "$APP_NAME" \
    --input target \
    --main-jar "$(basename "$JAR")" \
    --main-class com.example.CSVProcessorApp \
    --dest "$DEST" \
    --java-options "-Dfile.encoding=UTF-8" \
    --app-version 1.1.0 || true
fi
echo "Linux package(s) under: $DEST"
