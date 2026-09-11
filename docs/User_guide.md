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

1. Get a key at **[Google AI Studio](https://aistudio.google.com/apikey)**
2. Paste the whole key into the application and click **Save Key**

### Is it free?

Yes, for the model this app uses. `gemini-2.5-flash-lite` stays on the free
tier, with daily and per-minute request limits. If you hit them, the app
reports error 429 and you can continue the next day: progress is cached, so it
resumes from the last analyzed page. The **Pro** models left the free tier in
2026, but this app does not use them.

### The 2026 key format change

Keys issued today are *auth keys* and start with `AQ.`. The older keys started
with `AIza` and were exactly 39 characters long, and Google is retiring them.

* `AQ.` keys: current format, nothing to do.
* `AIza` keys: still accepted by the app, which shows a warning. Generate a new
  key when they stop working.

The app validates only that the key is not empty and has no spaces or line
breaks. It does not check the length, so a future format change will not lock
you out.

### Common errors

| Code | Meaning | Fix |
| :--- | :--- | :--- |
| 400 | The key was rejected | Check you copied it whole. If it starts with `AIza`, generate a new one. |
| 403 | The key was blocked | Usually a key published somewhere public, or a retired standard key. Generate a new one. |
| 429 | Daily or per-minute limit reached | Wait and resume. The cache keeps your progress. |

⚠️ Without a valid API Key, extraction will not work.

🔒 The key is saved locally in `src/api_key.env` and never leaves your machine.
Never commit that file or share the key.

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

For more help, check the documentation or open an *issue* in the repository.
