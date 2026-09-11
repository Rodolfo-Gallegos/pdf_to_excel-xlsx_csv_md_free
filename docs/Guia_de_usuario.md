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
* **Modelo**: utiliza **Gemini AI** (`gemini-2.5-flash-lite`) para máxima precisión
* **Persistencia y Caché**: guarda el progreso automáticamente; si el proceso se interrumpe, continúa desde la última página analizada
* **Formato Excel mejorado**: incluye separadores claros entre páginas para una mejor lectura de datos
* **Interfaz gráfica (GUI)** fácil de usar
* **Multi-formato**: exporta a Excel (`.xlsx`), CSV (`.csv`) y Markdown (`.md`)
* **Selección inteligente de páginas** usando lenguaje natural
* **Soporte multi-archivo** en una sola ejecución
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

1. Consigue tu clave en **[Google AI Studio](https://aistudio.google.com/apikey)**
2. Pega la clave completa en la aplicación y pulsa **Guardar Clave**

### ¿Sigue siendo gratis?

Sí, para el modelo que usa esta aplicación. `gemini-2.5-flash-lite` se mantiene
en la capa gratuita, con límites de peticiones por minuto y por día. Si los
alcanzas, la app muestra el error 429 y puedes continuar al día siguiente: el
progreso queda en caché y se reanuda desde la última página analizada. Lo que
salió de la capa gratuita en 2026 fueron los modelos **Pro**, que esta
aplicación no utiliza.

### El cambio de formato de 2026

Las claves que se generan hoy son *auth keys* y empiezan por `AQ.`. Las
anteriores empezaban por `AIza` y medían exactamente 39 caracteres, y Google
las está retirando.

* Claves `AQ.`: formato actual, no hay que hacer nada.
* Claves `AIza`: la app las sigue aceptando y te muestra un aviso. Genera una
  nueva cuando dejen de funcionar.

La aplicación solo comprueba que la clave no esté vacía y que no tenga espacios
ni saltos de línea. Ya no valida la longitud, así que un futuro cambio de
formato no te dejará fuera.

### Errores frecuentes

| Código | Significado | Solución |
| :--- | :--- | :--- |
| 400 | La clave fue rechazada | Comprueba que la copiaste completa. Si empieza por `AIza`, genera una nueva. |
| 403 | La clave fue bloqueada | Normalmente por publicarla en algún sitio público, o por ser una clave estándar ya retirada. Genera una nueva. |
| 429 | Límite diario o por minuto alcanzado | Espera y reanuda. La caché conserva tu progreso. |

⚠️ Sin una API Key válida, la extracción no funcionará.

🔒 La clave se guarda localmente en `src/api_key.env` y nunca sale de tu equipo.
Nunca subas ese archivo a un repositorio ni compartas la clave.

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
