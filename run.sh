#!/usr/bin/env bash
cd "$(dirname "$0")"

echo "============================================================"
echo "  CRM Pro  -  one-click setup and launch"
echo "============================================================"

# Check Python
if ! command -v python3 >/dev/null 2>&1; then
    echo "[ERROR] Python 3 not found. Install it and try again."
    exit 1
fi

# Auto-install dependencies if missing
echo "  Checking dependencies..."
if ! python3 -c "import flask, openpyxl" >/dev/null 2>&1; then
    echo "  First run - installing requirements (flask, openpyxl)..."
    python3 -m pip install --upgrade pip
    python3 -m pip install -r requirements.txt
fi

echo "  Starting CRM Pro..."
echo "  Your browser will open automatically at: http://127.0.0.1:5000"
echo "  Press Ctrl+C to stop the server."
echo "============================================================"
python3 app.py
