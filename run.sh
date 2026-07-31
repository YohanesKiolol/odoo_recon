#!/bin/bash
# run.sh — activates .venv and runs main.py
# Usage: ./run.sh [--bank bca] [--bank mandiri bri] etc.

set -e
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"

VENV="$SCRIPT_DIR/.venv/bin/python3"
if [ ! -f "$VENV" ]; then
  echo "ERROR: .venv not found. Run setup first:"
  echo "  python3 -m venv .venv"
  echo "  source .venv/bin/activate"
  echo "  pip install -r requirements.txt"
  exit 1
fi

"$VENV" "$SCRIPT_DIR/main.py" "$@"
