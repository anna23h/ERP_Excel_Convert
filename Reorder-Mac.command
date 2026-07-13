#!/bin/bash
# Reorder Helper - macOS double-click launcher
# First run: creates a local virtual environment (.venv) and installs deps; then launches.
cd "$(dirname "$0")" || exit 1

if ! command -v python3 >/dev/null 2>&1; then
  echo "python3 not found. Install Python 3 first: https://www.python.org/downloads/"
  read -n 1 -s -r -p "Press any key to close…"
  exit 1
fi

VENV=".venv"
if [ ! -d "$VENV" ]; then
  echo "First run: creating the environment and installing deps (about 1-2 minutes, only once)…"
  python3 -m venv "$VENV" || { echo "Failed to create virtual environment"; read -n 1 -s -r; exit 1; }
  "$VENV/bin/pip" install --upgrade pip >/dev/null 2>&1
  "$VENV/bin/pip" install -r requirements.txt || { echo "Failed to install dependencies"; read -n 1 -s -r; exit 1; }
fi

"$VENV/bin/python" reorder_gui.py
