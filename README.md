# PDF to EXCEL/CSV/MD AI extractor

An AI-powered tool that extracts tables from PDF files (scanned/digital) using Gemini AI.

## Result showcase

| 1. Source PDF | 2. Excel Result | 3. Markdown Result | 4. CSV Result |
| :---: | :---: | :---: | :---: |
| ![Original PDF](docs/screenshots/pdf_tables.png) | ![Excel Output](docs/screenshots/xlsx_table.png) | ![Markdown Output](docs/screenshots/markdown_table.png) | ![CSV Output](docs/screenshots/csv_table.png) |

## Interface showcase

| Prompt editor | Main menu | Extraction completed |
| :---: | :---: | :---: |
| ![Prompt editor](docs/screenshots/prompt_editor.png) | ![Main menu](docs/screenshots/before_extraction.png) | ![Extraction completed](docs/screenshots/extraction_completed.png) |

---

## How to run

### **Windows**

1. Double-click **`Windows_exec.bat`**.
2. It will automatically check Python, install dependencies, and create a desktop shortcut for you.

### **Linux & macOS**

1. Open terminal in this folder.
2. Run: `chmod +x Linux_exec.sh`
3. Run: `./Linux_exec.sh`

---

## Project structure

- `Windows_exec.bat`: Main launcher for Windows.
- `Linux_exec.sh`: Main launcher for Linux/macOS.
- `src/`: Source code and assets (Internal).
- `docs/`: Full documentation and screenshots.

## Documentation

- [Full documentation (English)](docs/User_guide.md)
- [Documentación completa (Español)](docs/Guia_de_usuario.md)

---

## Note on Excel results

The output Excel file contains a **"Summary"** sheet followed by a specific data sheet for each processed PDF file. You will find your tables starting from the second sheet.

## Selective processing

You can now ask the AI to process specific pages by modifying the prompt:

- _"Extract tables from page 2"_
- _"Extract tables from pages 1 to 3"_
