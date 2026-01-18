# PDF to EXCEL/CSV/MD AI Extractor

A robust AI-powered tool that extracts tables from PDF files by analyzing pages as images using the Gemini 3 Flash Preview model. Captures complex visual layouts that traditional text-based extractors fail to process.

## 📊 Result Showcase

### Input vs. Output

| 1. Source PDF | 2. Excel Result | 3. Markdown Result |
| :---: | :---: | :---: |
| ![Original PDF](screenshots/pdf_tables.png) | ![Excel Output](screenshots/xlsx_table.png) | ![Markdown Output](screenshots/markdown_table.png) |

> [!TIP]
> **From visual chaos to structured data in seconds.** Perfect for scanned documents, insurance quotes, and complex reports.

## ✨ Features

- **Multimodal AI**: Uses computer vision to "see" and extract tables exactly as they appear.
- **Premium GUI**: Modern, user-friendly interface with real-time logs and progress tracking.
- **Multi-format Export**: Save results to **Excel (.xlsx)**, **CSV**, and **Markdown**.
- **Data Cleaning**: One-click normalization to remove currency symbols and fix numeric formats.
- **Automated Setup**: One-click installer for Windows users.

## 🎥 Video Tutorial

[![Video Tutorial](https://img.shields.io/badge/YouTube-Video%20Tutorial-red?style=for-the-badge&logo=youtube)](https://www.youtube.com/watch?v=tu_video_id_aqui)
*In this video, I explain how to set up the repository and use both the GUI and CLI versions.*

---

## 🚀 Quick Start

### For Windows

1. Download or clone this repository.
2. Double-click **`setup_windows.bat`**.
   - *This will automatically install Python (if missing), setup dependencies, and launch the app.*

### For Linux (Ubuntu/Debian)

1. `sudo apt install python3-tk` (Optional: only needed for GUI).
2. `pip3 install -r requirements.txt`
3. Launch with `python3 gui_app.py` or use the CLI.

---

## 🛠️ How to Use

### Option 1: Graphical Interface (Recommended)

Launch the premium app to manage everything visually:

```bash
python3 gui_app.py
```

| Initial Setup | Extraction Progress |
| :---: | :---: |
| ![GUI Setup](screenshots/before_extraction.png) | ![GUI Progress](screenshots/extraction_completed.png) |

### Option 2: Command Line (Advanced/Automation)

Run the script directly for quick processing or automation:

```bash
python3 pdf_to_xlsx.py document.pdf --clean --md --csv -o final_results.xlsx
```

- `--clean`: Normalizes data (removes '$', ',', etc.).
- `--md` / `--csv`: Generates additional formats.

---

## ⚙️ Configuration & API Key

### 1. Requirements

- Python 3.8+
- A Google Gemini API Key

### 2. Setup your API Key

1. Get your free API key from [Google AI Studio](https://aistudio.google.com/api-keys).
2. Edit the existing `api_key.env` file in the root directory and replace the placeholder:

   ```env
   API_KEY=your_gemini_api_key_here
   ```

---

## 🏗️ Technical Details

1. **Rendering**: Uses `pdfplumber` to convert pages to 300 DPI images.
2. **Analysis**: Images are sent to **Gemini 3 Flash Preview** for table detection.
3. **Parsing**: AI Markdown is converted into `pandas` DataFrames.
4. **Writing**: Results are consolidated using `openpyxl`.

## 📜 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

---
*Bilingual Documentation: [English](README.md) | [Spanish](README_ES.md)*
