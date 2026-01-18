#!/bin/bash

# ==================================================
#   PDF Table Extractor - Unix Launcher (Linux/macOS)
# ==================================================

# Get the directory where the script is located
SCRIPT_DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" && pwd )"
cd "$SCRIPT_DIR"

echo "--------------------------------------------------"
echo "   PDF Table Extractor - Launching..."
echo "--------------------------------------------------"

# 1. Check if Python is installed
if ! command -v python3 &> /dev/null
then
    echo "[!] Python3 not found."
    if [[ "$OSTYPE" == "darwin"* ]]; then
        echo "Please install Python 3 using Homebrew: brew install python"
    else
        echo "Please install Python 3 using your package manager (apt, dnf, etc.)"
    fi
    exit 1
fi

# 2. Check if dependencies are installed
# We check for one of the main dependencies: google-genai
if ! python3 -c "import google.genai" &> /dev/null
then
    echo "[*] Dependencies missing. Installing from requirements.txt..."
    python3 -m pip install --upgrade pip
    python3 -m pip install -r requirements.txt
    if [ $? -ne 0 ]; then
        echo "[X] Failed to install dependencies. Check your internet connection or permissions."
        exit 1
    fi
    echo "[OK] Dependencies installed."
fi

# 3. Launch GUI
echo "[*] Starting Application..."
# Run in background and exit terminal
python3 gui_app.py &

echo "--------------------------------------------------"
echo "   Application started. You can close this window."
echo "--------------------------------------------------"
sleep 3
exit 0
