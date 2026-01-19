#!/bin/bash

# Get the directory where the script is located
SCRIPT_DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" && pwd )"
cd "$SCRIPT_DIR"

echo "--------------------------------------------------"
echo "   PDF Table Extractor - Launching..."
echo "--------------------------------------------------"

# Determine python command
if command -v python3 &>/dev/null; then
    PYTHON_CMD="python3"
elif command -v python &>/dev/null; then
    PYTHON_CMD="python"
else
    echo "Error: Python no está instalado / Python is not installed."
    exit 1
fi

# Install dependencies if missing
if ! $PYTHON_CMD -c "import google.genai" &> /dev/null
then
    echo "[*] Installing dependencies..."
    $PYTHON_CMD -m pip install --upgrade pip
    $PYTHON_CMD -m pip install -r src/requirements.txt
    if [ $? -ne 0 ]; then
        echo "[X] Failed to install dependencies."
        exit 1
    fi
    echo "[OK] Dependencies installed."
fi

# 3. Update Desktop Shortcut
bash "$SCRIPT_DIR/src/create_shortcut.sh"

# 4. Launch GUI
echo "[*] Starting Application..."
$PYTHON_CMD -m src.main &

echo "[OK] Application started."

echo "--------------------------------------------------"

echo "Right click on the desktop shortcut and select 'Allow Launching' to enable it."

echo "--------------------------------------------------"

sleep 2
exit 0
