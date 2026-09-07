#!/usr/bin/env bash
set -euo pipefail
ROOT="$(cd "$(dirname "$0")/.." && pwd)"
cd "$ROOT"

mvn -DskipTests package

APP_NAME="csv-to-excel-generator"
JAR="target/csv-to-excel-generator-1.3.0.jar"
STAGE="$ROOT/target/jpackage-input"
DEST="$ROOT/dist/linux"

rm -rf "$STAGE" "$DEST"
mkdir -p "$STAGE" "$DEST"
cp "$JAR" "$STAGE/"

jpackage \
  --type app-image \
  --name "$APP_NAME" \
  --input "$STAGE" \
  --main-jar "$(basename "$JAR")" \
  --main-class com.example.Launcher \
  --dest "$DEST" \
  --java-options "-Dfile.encoding=UTF-8" \
  --app-version 1.3.0

if command -v dpkg-deb >/dev/null 2>&1; then
  jpackage \
    --type deb \
    --name "$APP_NAME" \
    --input "$STAGE" \
    --main-jar "$(basename "$JAR")" \
    --main-class com.example.Launcher \
    --dest "$DEST" \
    --java-options "-Dfile.encoding=UTF-8" \
    --app-version 1.3.0 || true
fi

echo "Linux package(s) under: $DEST"
