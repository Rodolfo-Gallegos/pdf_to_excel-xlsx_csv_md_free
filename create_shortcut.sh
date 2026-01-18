#!/bin/bash

# ==================================================
#   Create Linux Desktop Shortcut (Button)
# ==================================================

SCRIPT_DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" && pwd )"
DESKTOP_FILE="$HOME/.local/share/applications/pdf_extractor.desktop"
ICON_PATH="$SCRIPT_DIR/icons/pdf_icon.png"

# Check if icon exists, if not use a generic one or none
if [ ! -f "$ICON_PATH" ]; then
    ICON_PATH="document-pdf"
fi

PYTHON_PATH=$(which python3)

echo "[*] Creating Desktop shortcut at $DESKTOP_FILE..."

cat <<EOF > "$DESKTOP_FILE"
[Desktop Entry]
Name=PDF Table Extractor
Comment=Extract tables from PDFs using Gemini AI
Exec=$PYTHON_PATH $SCRIPT_DIR/gui_app.py
Path=$SCRIPT_DIR
Icon=$ICON_PATH
Terminal=false
Type=Application
Categories=Office;Utility;
Keywords=pdf;excel;table;extractor;ai;
EOF

chmod +x "$DESKTOP_FILE"

echo "[OK] Shortcut created successfully!"
echo "     You can now find 'PDF Table Extractor' in your application menu."
