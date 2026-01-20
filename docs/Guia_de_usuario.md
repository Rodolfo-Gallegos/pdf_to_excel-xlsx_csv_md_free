# Extractor de PDF a EXCEL / CSV / MD con IA

Una **aplicación de escritorio** potenciada por IA que extrae tablas de archivos PDF (digitales o escaneados) y las convierte en **Excel, CSV o Markdown**, utilizando el modelo **Gemini**.

---

## Inicio rápido

1. En GitHub, haz clic en **<> Code → Download ZIP**
2. **Extrae** el archivo ZIP en una carpeta
3. Haz doble clic en **`Windows_exec.bat`**
4. Espera a que la aplicación se abra automáticamente
5. Pega tu **API Key de Gemini**
6. Haz clic en **Añadir archivos** y selecciona tus PDFs
7. Haz clic en **Iniciar extracción**

✅ Los archivos generados aparecerán en la carpeta **`extracted_tables`**.

---

## Ejemplo de resultados

|               1. PDF Original               |              2. Resultado Excel             |                3. Resultado Markdown               |             4. Resultado CSV             |
| :-----------------------------------------: | :-----------------------------------------: | :------------------------------------------------: | :--------------------------------------: |
| ![PDF Original](screenshots/pdf_tables.png) | ![Salida Excel](screenshots/xlsx_table.png) | ![Salida Markdown](screenshots/markdown_table.png) | ![Salida CSV](screenshots/csv_table.png) |

> 💡 **De PDF (incluso escaneado) a datos estructurados en segundos.** Ideal para reportes, estados de cuenta y documentos complejos.

---

## ✨ Características principales

* **IA multimodal**: análisis visual de páginas PDF como imágenes
* **Interfaz gráfica (GUI)** fácil de usar
* **Multi‑formato**: exporta a Excel (`.xlsx`), CSV (`.csv`) y Markdown (`.md`)
* **Selección inteligente de páginas** usando lenguaje natural
* **Soporte multi‑archivo** en una sola ejecución
* **Resultados organizados** con hoja de resumen en Excel

---

## Selección inteligente de páginas

El campo de *prompt* permite indicarle a la IA **qué páginas procesar** y **cómo hacerlo**, usando lenguaje natural en español o inglés.

### Selección básica

* **Página específica:** "Extraer tablas de la página 3"
* **Lista de páginas:** "Procesar páginas 1, 5 y 10"
* **Rango:** "Extraer de la página 2 a la 6"

### Selección por ordinales

El sistema entiende números ordinales:

* "Extraer la **primera** y la **última** página"
* "Procesar la **tercera** y **quinta** página"

Soporta ordinales en **español e inglés**.

### Filtrado por documento

Cuando se cargan varios PDFs:

* "Extraer página 1 de **ArchivoA.pdf** y la última de **ArchivoB.pdf**"
* "Extraer tablas solo de **Reporte_2024**"

---

## 📂 Estructura del proyecto

```text
PDF_to_XLSX/
├── Windows_exec.bat      # Lanzador principal para Windows
├── Linux_exec.sh         # Lanzador para Linux / macOS
├── README.md             # Guía rápida (Quick Start)
├── docs/                 # Documentación y capturas
│   ├── User_guide.md
│   └── Guia_de_usuario.md
└── src/                  # Código fuente (uso interno)
```

---

## Instalación y ejecución

### Windows (recomendado)

1. Descarga el proyecto como ZIP desde GitHub
2. **Extrae el ZIP** en una carpeta local
3. Haz doble clic en **`Windows_exec.bat`**

Durante la primera ejecución, el sistema:

* Verifica que Python esté instalado
* Instala automáticamente las dependencias
* Crea un **acceso directo en el escritorio**

⏳ La primera vez puede tardar **1 a 3 minutos**.

✅ Al finalizar, la aplicación se abrirá automáticamente.

---

### 🐧 Linux / 🍎 macOS

1. Abre una terminal en la carpeta del proyecto
2. Ejecuta:

   ```bash
   chmod +x Linux_exec.sh
   ```
3. Ejecuta:

   ```bash
   ./Linux_exec.sh
   ```

---

## Uso de la aplicación (GUI)

1. **Idioma**: Cambia entre Español / Inglés con el botón **EN / ES**
2. **API Key**: Pega tu clave de Gemini
3. **Prompt**: Usa el prompt por defecto o personalízalo
4. **Añadir archivos**: Selecciona uno o varios PDFs
5. **Ruta de salida**:

   * Por defecto: `extracted_tables/`
   * Puedes cambiarla si lo deseas
6. **Formato de salida**:

   * Excel (`.xlsx`)
   * CSV (`.csv`)
   * Markdown (`.md`)
7. Haz clic en **Iniciar extracción**

Al finalizar:

* Aparecerá un mensaje de confirmación
* Los archivos se guardarán en la ruta seleccionada

---

## 🔑 Configuración de la API Key

1. Obtén tu clave en **Google AI Studio (Gemini)**
2. Pégala directamente en la aplicación

⚠️ Sin una API Key válida, la extracción no funcionará.

---

## Resultados en Excel

* El archivo Excel generado contiene:

  * Una hoja **"Summary"** con el resumen general
  * Una hoja adicional por cada PDF procesado

---

## Detalles técnicos (para usuarios avanzados)

* Renderizado de páginas: `pdfplumber` (300 DPI)
* Procesamiento visual mediante Gemini
* La aplicación se ejecuta localmente; solo las imágenes se envían a la IA

---

Para más ayuda, consulta la documentación o abre un *issue* en el repositorio.
