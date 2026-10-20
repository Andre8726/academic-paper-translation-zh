#!/usr/bin/env bash
# Install PyMuPDF into a task-local directory for the macOS system Python.
# Usage: bootstrap_pymupdf.sh [TASK_DIR] [PYTHON]
set -euo pipefail

TASK_DIR="${1:-$PWD}"
PYTHON="${2:-/usr/bin/python3}"
DEST="$TASK_DIR/.pymupdf"
mkdir -p "$DEST" "$TASK_DIR/.pip-cache"

if PYTHONPATH="$DEST" "$PYTHON" -c 'import fitz' >/dev/null 2>&1; then
  echo "PyMuPDF already available: PYTHONPATH=$DEST $PYTHON"
  exit 0
fi

echo "Installing PyMuPDF 1.24.14 into $DEST ..."
PIP_CACHE_DIR="$TASK_DIR/.pip-cache" "$PYTHON" -m pip install \
  --disable-pip-version-check --no-input \
  --target "$DEST" "pymupdf==1.24.14"

PYTHONPATH="$DEST" "$PYTHON" -c 'import fitz; print("PyMuPDF", fitz.__doc__.splitlines()[0])'
echo "Use: PYTHONPATH=\"$DEST\" \"$PYTHON\" scripts/extract_pdf.py ..."
