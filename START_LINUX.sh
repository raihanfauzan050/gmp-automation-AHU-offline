#!/bin/bash
echo "============================================================"
echo "  GMP Automation System - Production Start"
echo "============================================================"
echo ""

# Check Python
if ! command -v python3 &> /dev/null; then
    echo "[ERROR] Python3 is not installed!"
    echo "Install with: sudo apt install python3 python3-pip"
    exit 1
fi

# Check poppler
if ! command -v pdftoppm &> /dev/null; then
    echo "[WARNING] Poppler is not installed!"
    echo "Install with: sudo apt install poppler-utils"
fi

# Install dependencies
echo "[1/2] Installing Python packages..."
python3 -m pip install -r requirements.txt

echo ""
echo "[2/2] Starting GMP Automation System with Gunicorn WSGI..."
echo ""
echo "============================================================"
echo "  Open your browser and go to: http://localhost:5002/offline"
echo "  Press Ctrl+C to stop the server."
echo "============================================================"
echo ""

exec python3 -m gunicorn --workers 1 --bind 0.0.0.0:5002 --timeout 600 app:app
