# PDF to XLSX Table Extractor

A robust Python tool that extracts tables from PDF files by analyzing pages as images using the Gemini 3 Flash Preview model. This is especially useful for PDFs where tables are embedded as images or have complex visual layouts that traditional text-based extractors fail to capture.

## Video Tutorial

[![Video Tutorial](https://img.shields.io/badge/YouTube-Video%20Tutorial-red?style=for-the-badge&logo=youtube)](https://www.youtube.com/watch?v=tu_video_id_aqui)
*En este video explico cómo usar la herramienta y cómo funciona el repositorio.*

## Features

- **Visual Extraction**: Uses Gemini's multimodal capabilities to identify and extract tables from PDF page renders.
- **Robust Parsing**: Converts AI-generated Markdown tables into clean pandas DataFrames.
- **Multiple Output Formats**:
  - **Excel**: Consolidated output in a single `.xlsx` file.
  - **Markdown**: Raw AI-generated tables saved as `.md`.
  - **CSV**: DataFrames exported as `.csv` files.
- **Data Normalization**: Optional `--clean` flag to remove currency symbols and standardize numeric formats for programmatic use.
- **Error Handling**: Implements retry logic for API quota limits and robust path validation.
- **Clean Output**: Silenced internal SDK logs for a cleaner terminal experience.
- **Bilingual Documentation**: Available in English and [Spanish (README_ES.md)](./README_ES.md).

## Visual Overview

### Extraction Flow

| 1. Source PDF | 2. GUI Setup | 3. Completion |
| :---: | :---: | :---: |
| ![Source PDF Tables](screenshots/pdf_tables.png) | ![GUI Configuration](screenshots/before_extraction.png) | ![Extraction Progress](screenshots/extraction_completed.png) |

*From source PDF to customized AI extraction in seconds.*

## Requirements

- Python 3.8+
- A Google Gemini API Key

## Installation

### Automated Setup (Recommended)

#### Windows

1. Download the project.
2. Double-click **`setup_windows.bat`**.
    - *This will automatically check for Python, install it if missing, setup dependencies, and launch the app.*

#### Linux (Ubuntu/Debian)

1. Install system dependencies:
   - **Optional (for GUI only)**:

     ```bash
     sudo apt install python3-tk
     ```

2. Install Python libraries:

   ```bash
   pip3 install -r requirements.txt
   ```

### Manual Installation (All Platforms)

1. Clone the repository:

   ```bash
   git clone https://github.com/your-username/pdf-to-xlsx-extractor.git
   cd pdf-to-xlsx-extractor
   ```

2. Install dependencies:

   ```bash
   pip3 install -r requirements.txt
   ```

3. Setup your API Key:
   - Get your free API key from [Google AI Studio](https://aistudio.google.com/api-keys).
   - Edit the existing `api_key.env` file in the root directory and replace the placeholder with your key:

   ```env
   API_KEY=your_gemini_api_key_here
   ```

## Usage

Run the script by providing the path to one or more PDF files:

```bash
python3 pdf_to_xlsx.py document1.pdf document2.pdf
```

### Advanced Options

- **Excel Output** (Default): `-o my_tables.xlsx`
- **Markdown Output**: `--md` (Saves a `.md` file for each PDF)
- **CSV Output**: `--csv` (Saves a `.csv` file for each PDF)
- **Data Cleaning**: `--clean` (Removes '$', ',', and normalizes numbers)

Example with all formats and cleaning:

```bash
python3 pdf_to_xlsx.py document.pdf --md --csv --clean -o final_results.xlsx
```

## Output Examples

| Excel (.xlsx) | Markdown (.md) | CSV (.csv) |
| :---: | :---: | :---: |
| ![Excel Result](screenshots/xlsx_table.png) | ![Markdown Result](screenshots/markdown_table.png) | ![CSV Result](screenshots/csv_table.png) |

## Graphical User Interface (GUI)

If you prefer a visual interface, you can use the `gui_app.py` script:

```bash
python3 gui_app.py
```

The GUI allows you to:

- Select multiple PDF files via a file browser.
- Configure output formats and data cleaning with checkboxes.
- Manage and save your Gemini API Key securely.
- Track progress through a real-time log area and progress bar.

## How it Works

1. **Page Rendering**: The script uses `pdfplumber` to convert each PDF page into a high-resolution image (300 DPI).
2. **AI Analysis**: Each image is sent to the Gemini 3 Flash Preview model with a prompt to extract all tables in Markdown format.
3. **Data Structuring**: The Markdown response is parsed into a pandas DataFrame.
4. **Excel Generation**: DataFrames are consolidated and written to an `.xlsx` file using `openpyxl`.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.
