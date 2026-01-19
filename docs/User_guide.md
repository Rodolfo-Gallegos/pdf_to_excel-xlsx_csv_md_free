# PDF to EXCEL/CSV/MD AI extractor

An AI-powered tool that extracts tables from PDF files by analyzing pages as images using the Gemini 3 Flash Preview model.

## 📊 Result showcase

| 1. Source PDF | 2. Excel Result | 3. Markdown Result | 4. CSV Result |
| :---: | :---: | :---: | :---: |
| ![Original PDF](screenshots/pdf_tables.png) | ![Excel Output](screenshots/xlsx_table.png) | ![Markdown Output](screenshots/markdown_table.png) | ![CSV Output](screenshots/csv_table.png) |

> [!TIP]
> **From pdf image to structured data in seconds.** Perfect for scanned documents and complex reports.

## ✨ Features

- **Multimodal AI**: Computer vision extraction.
- **Graphical Interface (GUI)**: Real-time logs and progress.
- **Multi-format Export**: Excel (.xlsx), CSV, and Markdown.
- **Selective Processing**: Ask for specific pages (e.g., "Page 2").

## 📂 Project Structure

```text
PDF_to_XLSX/
├── Windows_exec.bat     # Main Windows launcher
├── Linux_exec.sh        # Main Linux/macOS launcher
├── README.md            # Quick start guide
├── docs/                # Manuals and screenshots
│   ├── User_guide.md
│   └── Guia_de_usuario.md
└── src/                 # Source code and assets
    ├── assets/icons/    # Icon assets (pdf_to_excel.png)
    ├── ui/              # User Interface
    ├── logic/           # Processing logic
    ├── main.py          # GUI Entry point
    ├── cli.py           # CLI Entry point
    └── api_key.env      # API Key configuration
```

## 🚀 Quick Start

### For Windows

1. Double-click **`Windows_exec.bat`**.
2. It will automatically setup dependencies and create a desktop shortcut.

### For Linux & macOS

1. Open a terminal in the folder.
2. Run: `chmod +x Linux_exec.sh`
3. Run: `./Linux_exec.sh`
4. **Desktop Icon**: After running, you'll find the app in your menu. On Ubuntu, right-click the desktop icon and select **"Allow Launching"**.

---

## 🛠 How to Use

### Version 1: Graphical interface (GUI)

```bash
python -m src.main
```

### Version 2: Command Line (CLI)

```bash
python -m src.cli file1.pdf --output results.xlsx
```

## ⚙️ Configuration

1. Get your API key from [Google AI Studio](https://aistudio.google.com/api-keys).
2. Save it in the app or edit `src/api_key.env`.

## 🏗 Technical details

1. **Rendering**: `pdfplumber` (300 DPI).
2. **Analysis**: Gemini 3 Flash Preview.
3. **Consolidation**: Sheet "Summary" followed by data sheets.
