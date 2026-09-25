#!/usr/bin/env bash
set -euo pipefail
ROOT="$(cd "$(dirname "$0")" && pwd)"
cd "$ROOT"

PYTHON_BIN="${PYTHON:-python3}"
if ! command -v "$PYTHON_BIN" >/dev/null 2>&1; then
  PYTHON_BIN="python"
fi

echo "Starting PDF URL Capture with $PYTHON_BIN"
PORT="${PORT:-4747}" "$PYTHON_BIN" "$ROOT/launch.py"
