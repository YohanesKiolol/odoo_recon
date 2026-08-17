#!/bin/bash
set -e
echo "============================================================"
echo "  Building Standalone macOS SalesPortal.app"
echo "============================================================"
if [ -d ".venv" ]; then
    source .venv/bin/activate
fi
python3 build_sales.py
