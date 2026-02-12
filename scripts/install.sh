#!/usr/bin/env bash

set -euo pipefail

# Usage: ./setup_venv.sh [VENV_DIR]
# Default: .venv under pv_solar

VENV_DIR="${1:-.venv}"
PYTHON_BIN="${PYTHON:-python3}"

if ! command -v "${PYTHON_BIN}" >/dev/null 2>&1; then
  echo "python executable not found: ${PYTHON_BIN}" >&2
  exit 1
fi

"${PYTHON_BIN}" -m venv "${VENV_DIR}"
source "${VENV_DIR}/bin/activate"
python -m pip install --upgrade pip

# Runtime dependencies for pv_solar scripts
python -m pip install \
  python-pptx \
  pillow \
  openpyxl \
  reportlab \
  pymupdf \
  cairosvg

echo "Virtual environment ready at ${VENV_DIR}. Activate with: source ${VENV_DIR}/bin/activate"
