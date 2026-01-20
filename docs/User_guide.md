# AI-Powered PDF to EXCEL / CSV / MD Extractor

A **desktop application** powered by AI that extracts tables from PDF files (digital or scanned) and converts them into **Excel, CSV, or Markdown**, using the **Gemini** model.

---

## Quick Start

1. On GitHub, click **<> Code → Download ZIP**
2. **Extract** the ZIP file into a folder
3. Double-click **`Windows_exec.bat`**
4. Wait for the application to launch automatically
5. Paste your **Gemini API Key**
6. Click **Add files** and select your PDFs
7. Click **Start extraction**

✅ The generated files will appear in the **`extracted_tables`** folder.

---

## Example Results

|               1. Original PDF               |              2. Excel Output               |               3. Markdown Output              |              4. CSV Output               |
| :-----------------------------------------: | :----------------------------------------: | :-------------------------------------------: | :-------------------------------------: |
| ![Original PDF](screenshots/pdf_tables.png) | ![Excel Output](screenshots/xlsx_table.png) | ![Markdown Output](screenshots/markdown_table.png) | ![CSV Output](screenshots/csv_table.png) |

> 💡 **From PDF (even scanned) to structured data in seconds.** Ideal for reports, bank statements, and complex documents.

---

## ✨ Key Features

* **Multimodal AI**: visual analysis of PDF pages as images  
* **User-friendly GUI**
* **Multi-format export**: Excel (`.xlsx`), CSV (`.csv`), and Markdown (`.md`)
* **Smart page selection** using natural language
* **Multi-file support** in a single run
* **Organized results** with a summary sheet in Excel

---

## Smart Page Selection

The *prompt* field allows you to tell the AI **which pages to process** and **how**, using natural language in either Spanish or English.

### Basic Selection

* **Single page:** "Extract tables from page 3"
* **Page list:** "Process pages 1, 5, and 10"
* **Range:** "Extract from page 2 to 6"

### Ordinal Selection

The system understands ordinal numbers:

* "Extract the **first** and **last** page"
* "Process the **third** and **fifth** page"

Ordinal numbers are supported in **Spanish and English**.

### Document Filtering

When multiple PDFs are loaded:

* "Extract page 1 from **FileA.pdf** and the last page from **FileB.pdf**"
* "Extract tables only from **Report_2024**"

---

## 📂 Project Structure

```text
PDF_to_XLSX/
├── Windows_exec.bat      # Main launcher for Windows
├── Linux_exec.sh         # Launcher for Linux / macOS
├── README.md             # Quick Start guide
├── docs/                 # Documentation and screenshots
│   ├── User_guide.md
│   └── Guia_de_usuario.md
└── src/                  # Source code (internal use)
```

---

## Installation and Execution

### Windows (recommended)

1. Download the project as a ZIP from GitHub
2. **Extract the ZIP** into a local folder
3. Double-click **`Windows_exec.bat`**

During the first run, the system:

* Checks if Python is installed
* Automatically installs dependencies
* Creates a **desktop shortcut**

⏳ The first run may take **1 to 3 minutes**.

✅ Once completed, the application will launch automatically.

---

### 🐧 Linux / 🍎 macOS

1. Open a terminal in the project folder
2. Run:

   ```bash
   chmod +x Linux_exec.sh
   ```
3. Run:

   ```bash
   ./Linux_exec.sh
   ```

---

## Application Usage (GUI)

1. **Language**: Switch between Spanish / English using the **EN / ES** button
2. **API Key**: Paste your Gemini API key
3. **Prompt**: Use the default prompt or customize it
4. **Add files**: Select one or more PDF files
5. **Output path**:
   * Default: `extracted_tables/`
   * Can be changed if desired
6. **Output format**:
   * Excel (`.xlsx`)
   * CSV (`.csv`)
   * Markdown (`.md`)
7. Click **Start extraction**

After completion:

* A confirmation message will appear
* Files will be saved in the selected output path

---

## 🔑 API Key Configuration

1. Obtain your key from **Google AI Studio (Gemini)**
2. Paste it directly into the application

⚠️ Without a valid API Key, extraction will not work.

---

## Excel Output Details

* The generated Excel file contains:
  * A **"Summary"** sheet with a general overview
  * One additional sheet per processed PDF

---

## Technical Details (Advanced Users)

* Page rendering: `pdfplumber` (300 DPI)
* Visual processing via Gemini
* The application runs locally; only page images are sent to the AI

---

## ❓ Common Issues

* **App does not open** → Make sure the ZIP was extracted
* **No files generated** → Verify that the API Key is valid
* **PDF has no tables** → The document may not contain detectable tables

---

For more help, check the documentation or open an *issue* in the repository.
