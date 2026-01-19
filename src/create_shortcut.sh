#!/bin/bash

# ==================================================
# Get the project root (one level up from src/)
SRC_DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" && pwd )"
ROOT_DIR="$( dirname "$SRC_DIR" )"
DESKTOP_FILE_NAME="pdf_extractor.desktop"
APPLICATIONS_DIR="$HOME/.local/share/applications"

# Support localized Desktop folder names
DESKTOP_DIR="$HOME/Desktop"
if [ ! -d "$DESKTOP_DIR" ]; then
    DESKTOP_DIR="$HOME/Escritorio"
fi

ICON_PATH="$SRC_DIR/assets/icons/pdf_to_excel.png"

# Check if icon exists, if not use a generic one
if [ ! -f "$ICON_PATH" ]; then
    ICON_PATH="document-pdf"
fi

PYTHON_PATH=$(which python3)
if [ -z "$PYTHON_PATH" ]; then
    PYTHON_PATH=$(which python)
fi

echo "[*] Updating shortcuts..."

# Content of the desktop entry
ENTRY_CONTENT="[Desktop Entry]
Name=PDF Table Extractor
Comment=Extract tables from PDFs using Gemini AI
Exec=$PYTHON_PATH -m src.main
Path=$ROOT_DIR
Icon=$ICON_PATH
Terminal=false
Type=Application
Categories=Office;Utility;
Keywords=pdf;excel;table;extractor;ai;"

# 1. Update Menu Shortcut
echo "[*] Updating Applications Menu shortcut..."
mkdir -p "$APPLICATIONS_DIR"
echo "$ENTRY_CONTENT" > "$APPLICATIONS_DIR/$DESKTOP_FILE_NAME"
chmod +x "$APPLICATIONS_DIR/$DESKTOP_FILE_NAME"

# 2. Update Desktop Shortcut (if Desktop folder exists)
if [ -d "$DESKTOP_DIR" ]; then
    echo "[*] Creating shortcut on Desktop..."
    echo "$ENTRY_CONTENT" > "$DESKTOP_DIR/$DESKTOP_FILE_NAME"
    chmod +x "$DESKTOP_DIR/$DESKTOP_FILE_NAME"
    # On GNOME (Ubuntu), we might need to tell the user to "Allow Launching"
fi

echo "[OK] Shortcuts updated successfully!"
echo "     - Check your application menu for 'PDF Table Extractor'."
if [ -d "$DESKTOP_DIR" ]; then
    echo "     - Check your Desktop screen (Right-click -> 'Allow Launching' if it shows a gear icon)."
fi
