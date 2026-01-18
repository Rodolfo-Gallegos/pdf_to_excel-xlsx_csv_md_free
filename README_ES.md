# Extractor de PDF a EXCEL/CSV/MD con IA

Una herramienta potenciada por IA que extrae tablas de archivos PDF analizando las páginas como imágenes mediante el modelo Gemini 3 Flash Preview. Captura diseños visuales complejos que los extractores basados en texto tradicional no logran procesar.

_Documentación: [English](README.md) | [Español](README_ES.md)_

## 📊 Resumen de Resultados

### PDF Origen vs. Salida

| 1. PDF Original | 2. Resultado Excel | 3. Resultado Markdown | 4. Resultado CSV |
| :---: | :---: | :---: | :---: |
| ![PDF Original](screenshots/pdf_tables.png) | ![Salida Excel](screenshots/xlsx_table.png) | ![Salida Markdown](screenshots/markdown_table.png) | ![Salida CSV](screenshots/csv_table.png) |

> [!TIP]
> **De imagen en pdf a datos estructurados en segundos.** Ideal para documentos escaneados y reportes complejos.

## ✨ Características

- **IA Multimodal**: Utiliza visión artificial para extraer tablas exactamente como aparecen.
- **Interfaz Gráfica (GUI)**: Pantalla de uso sencillo con registro en tiempo real y barra de progreso.
- **Exportación Multi-formato**: Guarda resultados en **Excel (.xlsx)**, **CSV** y **Markdown**.
- **Limpieza de Datos**: Normalización opcional para eliminar símbolos de moneda y fijar formatos numéricos.
- **Instalación Automatizada**: Script de configuración para usuarios de Windows.

## 🤖 Compatibilidad con LLMs y Propósito

El objetivo principal de este proyecto es proporcionar **datos estructurados y limpios** a partir de PDFs complejos. Los formatos generados (especialmente Markdown y CSV) están diseñados para ser "copiados y pegados" fácilmente en otros Modelos de Lenguaje (ChatGPT, Claude, Gemini, etc.).

Esto permite aprovechar la IA para tareas adicionales como:

- Convertir los datos a tablas en **LaTeX**.
- Realizar una **limpieza profunda** o análisis de datos avanzado.
- Reformatear los resultados en estructuras personalizadas de JSON o código.

## 🧠 Personalización de IA (Ajuste del Prompt)

Puedes personalizar cómo la IA analiza cada página modificando el **prompt del sistema**. Esto es útil si deseas extraer texto que no sean tablas, solicitar formatos especializados (como **LaTeX**) o aplicar lógica personalizada.

El prompt se encuentra en:

- **CLI**: `pdf_to_xlsx.py` dentro de la función `process_page_images()`.
- **GUI**: `gui_app.py` dentro del método `_extract_from_page()`.

### Prompt Actual
>
> _"Analyze this page and extract ALL tables you see. Even if the table looks like a screenshot or an embedded image, extract it. Return results strictly in Markdown format. Do not include any introductory text, titles outside the table, or comments. If no tables are found, return an empty string."_

Al modificar estas líneas, puedes decirle a Gemini:

- _"Extrae todos los encabezados y la tabla principal..."_
- _"Formatea el resultado como una tabla larga (longtable) de LaTeX..."_
- _"Resume el texto sobre la tabla y luego extrae los datos..."_

<!-- [![Video Tutorial](https://img.shields.io/badge/YouTube-Video%20Tutorial-red?style=for-the-badge&logo=youtube)](https://www.youtube.com/watch?v=tu_video_id_aqui)
_En este video explico cómo configurar el repositorio y cómo usar tanto la versión gráfica como la de terminal._ -->

---

## 🚀 Inicio Rápido

### Para Windows

1. Descarga o clona este repositorio.
2. Haz doble clic en **`setup_windows.bat`**.
   - _Esto instalará Python automáticamente (si falta), configurará las dependencias e iniciará la app._

### Para Linux (Ubuntu/Debian)

1. `sudo apt install python3-tk` (Opcional: solo necesario para la GUI).
2. `pip3 install -r requirements.txt`
3. Inicia con `python3 gui_app.py` o usa la terminal (CLI).

---

## 🛠️ Modo de Uso

### Opción 1: Interfaz Gráfica (Recomendado)

Inicia la aplicación para gestionar todo visualmente:

```bash
python3 gui_app.py
```

| Configuración Inicial | Progreso de Extracción |
| :---: | :---: |
| ![Setup GUI](screenshots/before_extraction.png) | ![Progreso GUI](screenshots/extraction_completed.png) |

### Opción 2: Línea de Comandos (Avanzado/Automatización)

Ejecuta el script directamente para procesamiento rápido o automatización:

```bash
python3 pdf_to_xlsx.py documento.pdf --clean --md --csv -o tablas_excel.xlsx
```

- `--clean`: Normaliza los datos (quita '$', ',', etc.).
- `--md` / `--csv`: Genera formatos adicionales.

---

## ⚙️ Configuración y Llave API

### 1. Requisitos

- Python 3.8+
- Una llave de API de Google Gemini

### 2. Configura tu API Key

1. Consigue tu clave gratuita en [Google AI Studio](https://aistudio.google.com/api-keys).
2. Edita el archivo `api_key.env` existente en la raíz y sustituye el marcador:

   ```env
   API_KEY=tu_clave_de_api_gemini_aqui
   ```

---

## 🏗️ Detalles Técnicos

1. **Renderizado**: Usa `pdfplumber` para convertir páginas en imágenes de 300 DPI.
2. **Análisis**: Las imágenes se envían a **Gemini 3 Flash Preview** para detectar tablas.
3. **Procesamiento**: El Markdown de la IA se convierte en DataFrames de `pandas`.
4. **Escritura**: Los resultados se consolidan usando `openpyxl`.

## 📜 Licencia

Este proyecto está bajo la Licencia MIT - mira el archivo [LICENSE](LICENSE) para más detalles.

---

