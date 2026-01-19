# Extractor de PDF a EXCEL/CSV/MD con IA

Una herramienta potenciada por IA que extrae tablas de archivos PDF analizando las páginas como imágenes mediante el modelo Gemini 3 Flash Preview.

## 📊 Resumen de resultados

| 1. PDF Original | 2. Resultado Excel | 3. Resultado Markdown | 4. Resultado CSV |
| :---: | :---: | :---: | :---: |
| ![PDF Original](screenshots/pdf_tables.png) | ![Salida Excel](screenshots/xlsx_table.png) | ![Salida Markdown](screenshots/markdown_table.png) | ![Salida CSV](screenshots/csv_table.png) |

> [!TIP]
> **De imagen en pdf a datos estructurados en segundos.** Ideal para documentos escaneados y reportes complejos.

## ✨ Características

- **IA Multimodal**: Extracción mediante visión artificial.
- **Interfaz Gráfica (GUI)**: Registro en tiempo real y progreso.
- **Multi-formato**: Excel (.xlsx), CSV y Markdown.
- **Procesamiento Selectivo**: Control avanzado sobre qué páginas analizar mediante lenguaje natural.

---

## 🧠 Selección inteligente de páginas

El "prompt" de la IA no solo sirve para decirle a Gemini cómo extraer los datos, sino también para especificar **qué** datos mirar. Puedes usar lenguaje natural para filtrar páginas y documentos.

### Selección básica

- **Página única:** _"Extraer tablas de la página 3"_
- **Listas:** _"Procesar páginas 1, 5 y 10"_
- **Rangos:** _"Obtener datos usando las páginas 2 a la 6"_

### Selección por ordinales (Palabras clave)

El sistema entiende números ordinales (tanto en español como en inglés):

- _"Extraer la **primera** página y la **última** página"_
- _"Procesar la **tercera** y **quinta** página"_
- **Palabras soportadas:** primera, segunda, ..., décima, última (y sus variantes en inglés).

### Filtrado por documento

Al procesar varios archivos a la vez, puedes dirigir la instrucción a archivos específicos:

- _"Extraer página 1 de **Archivo_A.pdf** y la última página de **Archivo_B.pdf**"_
- _"Extraer tablas de **Doc1**"_ (Esto omitirá otros archivos en la cola de procesamiento)

---

## 📂 Estructura del proyecto

```text
PDF_to_XLSX/
├── Windows_exec.bat     # Lanzador principal Windows
├── Linux_exec.sh        # Lanzador principal Linux/macOS
├── README.md            # Guía rápida
├── docs/                # Manuales y capturas
│   ├── User_guide.md
│   └── Guia_de_usuario.md
└── src/                 # Código fuente y activos
    ├── assets/icons/    # Iconos (pdf_to_excel.png)
    ├── ui/              # Interfaz
    ├── logic/           # Lógica de procesamiento
    ├── main.py          # Punto de entrada GUI
    ├── cli.py           # Punto de entrada CLI
    └── api_key.env      # Configuración de Clave API
```

## 🚀 Inicio rápido

### En Windows

1. Haz doble clic en **`Windows_exec.bat`**.
2. Instalará dependencias y creará un acceso directo en el escritorio.

### En Linux & macOS

1. Abre una terminal en la carpeta.
2. Ejecuta: `chmod +x Linux_exec.sh`
3. Ejecuta: `./Linux_exec.sh`
4. **Icono de Escritorio**: Tras la primera ejecución, aparecerá en tu menú. En Ubuntu, haz clic derecho en el icono del escritorio y elige **"Permitir lanzar"**.

---

## 🛠 Modo de uso

### Versión 1: Interfaz Gráfica (GUI)

```bash
python -m src.main
```

### Versión 2: Línea de Comandos (CLI)

```bash
python -m src.cli archivo.pdf --output resultados.xlsx
```

## ⚙️ Configuración

1. Consigue tu clave en [Google AI Studio](https://aistudio.google.com/api-keys).
2. Guárdala en la app o edita `src/api_key.env`.

## 🏗 Detalles técnicos

1. **Renderizado**: `pdfplumber` (300 DPI).
2. **Consolidación**: Hoja "Summary" seguida de hojas de datos por archivo.
