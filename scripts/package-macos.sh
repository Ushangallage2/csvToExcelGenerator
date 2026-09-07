#!/usr/bin/env bash
set -euo pipefail
ROOT="$(cd "$(dirname "$0")/.." && pwd)"
cd "$ROOT"

export JAVA_HOME="${JAVA_HOME:-$HOME/.sdkman/candidates/java/21.0.6-tem}"
export PATH="$JAVA_HOME/bin:$PATH"

mvn -DskipTests package

APP_NAME="CSV to Excel Generator"
JAR="target/csv-to-excel-generator-1.3.0.jar"
STAGE="$ROOT/target/jpackage-input"
DEST="$ROOT/dist/macos"

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

APP_PATH="$DEST/$APP_NAME.app"
xattr -cr "$APP_PATH" 2>/dev/null || true

echo "macOS app image: $APP_PATH"
echo "Launch with: open \"$APP_PATH\""
