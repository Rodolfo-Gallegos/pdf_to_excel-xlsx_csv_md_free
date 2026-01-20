# PDF to EXCEL / CSV / MD AI Extractor

An **AI-powered desktop application** that extracts tables from PDF files (scanned or digital) and exports them to **Excel, CSV, or Markdown** using **Gemini AI**.

---

## Quick start

1. Click **<> Code → Download ZIP**
2. Extract the ZIP file
3. Double-click **`Windows_exec.bat`**
4. Wait for the application to open
5. Paste your **Gemini API key**
6. Click **Add files** and select your PDF files
7. Click **Start extraction**

✅ Extracted tables will be saved in the **`extracted_tables`** folder.

---

## What this tool does

* Extracts tables from one or multiple PDF files
* Works with **scanned and digital PDFs**
* Exports results to:
  * Excel (`.xlsx`)
  * CSV (`.csv`)
  * Markdown (`.md`)
* Allows smart page selection using natural language prompts
* Supports **English and Spanish** interfaces

---

## Result showcase

|                   1. Source PDF                  |                  2. Excel Result                 |                    3. Markdown Result                   |                 4. CSV Result                 |
| :----------------------------------------------: | :----------------------------------------------: | :-----------------------------------------------------: | :-------------------------------------------: |
| ![Original PDF](docs/screenshots/pdf_tables.png) | ![Excel Output](docs/screenshots/xlsx_table.png) | ![Markdown Output](docs/screenshots/markdown_table.png) | ![CSV Output](docs/screenshots/csv_table.png) |

---

## Interface showcase

|                     Prompt editor                    |                       Main menu                      |                        Extraction completed                        |
| :--------------------------------------------------: | :--------------------------------------------------: | :----------------------------------------------------------------: |
| ![Prompt editor](docs/screenshots/prompt_editor.png) | ![Main menu](docs/screenshots/before_extraction.png) | ![Extraction completed](docs/screenshots/extraction_completed.png) |

---

## Step-by-step guide

### 1. Download the project

1. Click the **<> Code** button on this GitHub page
2. Select **Download ZIP**
3. Wait for the download to finish

---

### 2. Extract the ZIP file

1. Right‑click the downloaded ZIP file
2. Click **Extract All** (Windows) or **Extract here**
3. Open the extracted folder

⚠️ **Important:** Do NOT run the program from inside the ZIP file.

---

### 3. Run the application

#### **Windows (recommended)**

1. Double‑click **`Windows_exec.bat`**
2. The program will automatically:

   * Check if Python is installed
   * Install all required dependencies
   * Create a **desktop shortcut**

⏳ The first run may take **1–3 minutes**. Please be patient.

✅ When finished, the application will **open automatically**.

---

#### **🐧 Linux / 🍎 macOS**

1. Open a terminal in the extracted folder
2. Run:

   ```bash
   chmod +x Linux_exec.sh
   ```
3. Run:

   ```bash
   ./Linux_exec.sh
   ```

---

## Using the application

### 4. Change language (optional)

* Click the **EN / ES** button to switch between English and Spanish

---

### 5. Paste your API key 🔑

1. Copy your **Gemini API key**
2. Paste it into the **API Key** field in the app

You can obtain an API key from:

* Google AI Studio (Gemini)

⚠️ Without an API key, extraction will not work.

---

### 6. Configure the prompt (optional)

* You may **use the default prompt** (recommended for beginners)
* Or customize it to control:

  * Pages to extract
  * Specific files
  * Extraction behavior

---

### 7. Add PDF files

1. Click **Add files**
2. Select one or more PDF files
3. Confirm your selection

---

### 8. Configure output

* **Output folder**:

  * A default folder is created automatically:

    ```
    extracted_tables/
    ```
  * You may change the output path if desired

* **File formats**:

  * Choose one or more:

    * Excel (`.xlsx`)
    * CSV (`.csv`)
    * Markdown (`.md`)

* **File names**:

  * Customize output file names if needed

---

### 9. Start extraction

1. Review your settings
2. Click **Start extraction**
3. Wait for processing to finish

✅ When completed:

* A success message will appear
* Your files will be saved in the selected output folder

---

## Smart page selection (advanced)

You can control which pages or files are processed by editing the prompt.

### Examples (English)

* "Extract pages 1, 3 and 5"
* "Extract from page 2 to 4"
* "Extract the first page and the last page"
* "Extract page 1 from DocumentA and page 2 from DocumentB"

### Ejemplos (Español)

* "Extraer páginas 1, 3 y 5"
* "Extraer de la página 2 a la 4"
* "Extraer la primera y la última página"
* "Extraer página 1 de ArchivoA y página 2 de ArchivoB"

---

## Project structure

* `Windows_exec.bat` – Windows launcher (recommended)
* `Linux_exec.sh` – Linux/macOS launcher
* `src/` – Internal source code
* `docs/` – Documentation and screenshots

---

## Documentation

* [Full documentation (English)](docs/User_guide.md)
* [Documentación completa (Español)](docs/Guia_de_usuario.md)

---

## Notes

* The Excel output contains a **Summary** sheet followed by one sheet per processed PDF
* The application runs **locally**; only the AI request is sent to Gemini
* First run is slower due to dependency installation

---

If you encounter issues, please check the documentation or open an issue on GitHub.
