# Extractor de Tablas de PDF a XLSX

Una herramienta robusta en Python que extrae tablas de archivos PDF analizando las páginas como imágenes mediante el modelo Gemini 3 Flash Preview. Esto es especialmente útil para PDFs donde las tablas están incrustadas como imágenes o tienen diseños visuales complejos que los extractores basados en texto tradicional no logran capturar.

## Video Tutorial

[![Video Tutorial](https://img.shields.io/badge/YouTube-Video%20Tutorial-red?style=for-the-badge&logo=youtube)](https://www.youtube.com/watch?v=tu_video_id_aqui)
*En este video explico cómo usar la herramienta y cómo funciona el repositorio.*

## Características

- **Extracción Visual**: Utiliza las capacidades multimodales de Gemini para identificar y extraer tablas a partir de renders de páginas PDF.
- **Procesamiento Robusto**: Convierte tablas Markdown generadas por IA en DataFrames de pandas limpios.
- **Múltiples Formatos de Salida**:
  - **Excel**: Salida consolidada en un único archivo `.xlsx`.
  - **Markdown**: Tablas originales generadas por IA guardadas como `.md`.
  - **CSV**: DataFrames exportados como archivos `.csv`.
- **Normalización de Datos**: Bandera opcional `--clean` para eliminar símbolos de moneda y estandarizar formatos numéricos para uso programático.
- **Gestión de Errores**: Implementa lógica de reintento para límites de cuota de API y validación robusta de rutas.
- **Salida Limpia**: Se silenciaron los logs internos del SDK para una experiencia en terminal más limpia.
- **Documentación Bilingüe**: Disponible en Inglés y [Español (README_ES.md)](./README_ES.md).

## Resumen Visual

### Flujo de Extracción

| 1. PDF Origen | 2. Configuración GUI | 3. Finalización |
| :---: | :---: | :---: |
| ![Tablas PDF Origen](screenshots/pdf_tables.png) | ![Configuración GUI](screenshots/before_extraction.png) | ![Progreso de Extracción](screenshots/extraction_completed.png) |

*Desde el PDF origen hasta una extracción personalizada por IA en segundos.*

## Requisitos

- Python 3.8+
- Una clave de API de Google Gemini

## Instalación

### Instalación Automatizada (Recomendado)

#### Windows

1. Descarga el proyecto.
2. Haz doble clic en **`setup_windows.bat`**.
    - *Esto buscará Python automáticamente, lo instalará si falta, configurará las dependencias e iniciará la aplicación.*

#### Linux (Ubuntu/Debian)

1. Instala las dependencias del sistema:
   - **Opcional (solo para la GUI)**:

     ```bash
     sudo apt install python3-tk
     ```

2. Instala las librerías de Python:

   ```bash
   pip3 install -r requirements.txt
   ```

### Instalación Manual (Todas las plataformas)

1. Clona el repositorio:

   ```bash
   git clone https://github.com/tu-usuario/pdf-to-xlsx-extractor.git
   cd pdf-to-xlsx-extractor
   ```

2. Instala las dependencias:

   ```bash
   pip3 install -r requirements.txt
   ```

3. Configura tu clave de API:
   - Consigue tu clave de API gratuita en [Google AI Studio](https://aistudio.google.com/api-keys).
   - Edita el archivo `api_key.env` existente en el directorio raíz y sustituye el valor por tu clave:

   ```env
   API_KEY=tu_clave_de_api_gemini_aqui
   ```

## Uso

Ejecuta el script proporcionando la ruta de uno o más archivos PDF:

```bash
python3 pdf_to_xlsx.py documento1.pdf documento2.pdf
```

### Opciones Avanzadas

- **Salida Excel** (Predeterminado): `-o mis_tablas.xlsx`
- **Salida Markdown**: `--md` (Guarda un archivo `.md` por cada PDF)
- **Salida CSV**: `--csv` (Guarda un archivo `.csv` por cada PDF)
- **Limpieza de Datos**: `--clean` (Elimina '$', ',' y normaliza números)

Ejemplo con todos los formatos y limpieza:

```bash
python3 pdf_to_xlsx.py documento.pdf --md --csv --clean -o resultados_finales.xlsx
```

## Ejemplos de Salida

| Excel (.xlsx) | Markdown (.md) | CSV (.csv) |
| :---: | :---: | :---: |
| ![Resultado Excel](screenshots/xlsx_table.png) | ![Resultado Markdown](screenshots/markdown_table.png) | ![Resultado CSV](screenshots/csv_table.png) |

## Interfaz Gráfica (GUI)

Si prefieres una interfaz visual, puedes usar el script `gui_app.py`:

```bash
python3 gui_app.py
```

La GUI te permite:

- Seleccionar múltiples archivos PDF mediante un explorador de archivos.
- Configurar los formatos de salida y la limpieza de datos con casillas de verificación.
- Gestionar y guardar tu clave de API Gemini de forma segura.
- Seguir el progreso a través de un área de registro en tiempo real y una barra de progreso.

## Cómo funciona

1. **Renderizado de Página**: El script utiliza `pdfplumber` para convertir cada página del PDF en una imagen de alta resolución (300 DPI).
2. **Análisis de IA**: Cada imagen se envía al modelo Gemini 3 Flash Preview con un prompt para extraer todas las tablas en formato Markdown.
3. **Estructuración de Datos**: La respuesta en Markdown se procesa en un DataFrame de pandas.
4. **Generación de Excel**: Los DataFrames se consolidan y se escriben en un archivo `.xlsx` usando `openpyxl`.

## Licencia

Este proyecto está bajo la Licencia MIT - mira el archivo [LICENSE](LICENSE) para más detalles.
