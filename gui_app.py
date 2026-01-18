import os
import sys
import time
import logging
import threading
import webbrowser
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk
from typing import List, Optional

import pdfplumber
import pandas as pd
from google import genai
from dotenv import load_dotenv, set_key
from PIL import Image, ImageTk

# Metadata for versioning/help
VERSION = "1.4.0"

DEFAULT_PROMPT = """
Analyze this page and extract ALL tables you see.
Even if the table looks like a screenshot or an embedded image, extract it.
Return results strictly in Markdown format.
Do not include any introductory text, titles outside the table, or comments.
If no tables are found, return an empty string.
"""

TEXTS = {
    "EN": {
        "title": "PDF to EXCEL/CSV/MD AI Extractor",
        "config_section": " 1. CONFIGURATION ",
        "api_key": "API Key:",
        "save_key": "Save Key",
        "get_key": "Get your free API key at Google AI Studio",
        "pdf_selection": " 2. PDF SELECTION ",
        "select_files": "Select files to process:",
        "add_files": "+ Add Files",
        "clear": "🗑 Clear",
        "output_config": " 3. OUTPUT CONFIGURATION ",
        "output_folder": "Output Folder:",
        "browse": "Browse...",
        "excel_name": "Excel Name:",
        "csv_name": "CSV Name:",
        "md_name": "MD Name:",
        "options_section": " 4. OPTIONS ",
        "opt_excel": "Excel (.xlsx)",
        "opt_md": "Markdown (.md)",
        "opt_csv": "CSV (.csv)",
        "opt_normalize": "Normalize Data",
        "start_btn": " START EXTRACTION ",
        "status_log": " STATUS LOG ",
        "edit_prompt": "Edit Prompt",
        "ilovepdf": "Other PDF tools (iLovePDF)",
        "language": "Language:",
        "save": "Save",
        "reset": "Reset",
        "cancel": "Cancel",
        "prompt_editor_title": "Prompt Editor",
        "success": "Success",
        "warning": "Warning",
        "error": "Error",
        "key_saved": "API Key saved to api_key.env",
        "no_key": "Gemini API Key is required.",
        "key_length": "Incorrect API Key length. The key must be 39 characters long.",
        "no_files": "Please add at least one PDF file.",
        "process_finished": "Process Finished",
        "process_success": "The extraction process has completed successfully!",
        "process_error": "Process finished with errors.",
        "all_tasks_done": "SYSTEM: All tasks completed.",
        "analyzing_page": "    - Analyzing page {}...",
        "working_on": "Working on: {}",
        "done": "DONE: {}",
        "skip": "SKIP: No tables in {}",
        "saved_md": "  + Saved MD: {}",
        "saved_csv": "  + Saved CSV: {}",
        "files_added": "Added {} new files.",
        "files_cleared": "File selection cleared.",
        "output_path_set": "Output path: {}",
        "fatal_error": "The process encountered a fatal error",
        "quota_error": "The daily API limit has been exceeded (Code 429).",
        "api_error": "The API key entered is incorrect (Code 400). Please verify it and try again."
    },
    "ES": {
        "title": "Extractor de Tablas PDF con IA",
        "config_section": " 1. CONFIGURACIÓN ",
        "api_key": "Clave API:",
        "save_key": "Guardar",
        "get_key": "Obtén tu clave API gratis en Google AI Studio",
        "pdf_selection": " 2. SELECCIÓN DE PDF ",
        "select_files": "Selecciona archivos para procesar:",
        "add_files": "+ Añadir Archivos",
        "clear": "🗑 Limpiar",
        "output_config": " 3. CONFIGURACIÓN DE SALIDA ",
        "output_folder": "Carpeta de Salida:",
        "browse": "Buscar...",
        "excel_name": "Nombre Excel:",
        "csv_name": "Nombre CSV:",
        "md_name": "Nombre MD:",
        "options_section": " 4. OPCIONES ",
        "opt_excel": "Excel (.xlsx)",
        "opt_md": "Markdown (.md)",
        "opt_csv": "CSV (.csv)",
        "opt_normalize": "Normalizar Datos",
        "start_btn": " INICIAR EXTRACCIÓN ",
        "status_log": " REGISTRO DE ESTADO ",
        "edit_prompt": "Editar Prompt",
        "ilovepdf": "Otras herramientas PDF (iLovePDF)",
        "language": "Idioma:",
        "save": "Guardar",
        "reset": "Reiniciar",
        "cancel": "Cancelar",
        "prompt_editor_title": "Editor de Prompt",
        "success": "Éxito",
        "warning": "Advertencia",
        "error": "Error",
        "key_saved": "Clave API guardada en api_key.env",
        "no_key": "Se requiere la clave API de Gemini.",
        "key_length": "Longitud de API incorrecta. La clave debe tener 39 caracteres.",
        "no_files": "Por favor, añade al menos un archivo PDF.",
        "process_finished": "Proceso Finalizado",
        "process_success": "¡El proceso de extracción ha finalizado con éxito!",
        "process_error": "Proceso finalizado con errores.",
        "all_tasks_done": "SISTEMA: Todas las tareas completadas.",
        "analyzing_page": "    - Analizando página {}...",
        "working_on": "Trabajando en: {}",
        "done": "LISTO: {}",
        "skip": "OMITIR: No hay tablas en {}",
        "saved_md": "  + MD Guardado: {}",
        "saved_csv": "  + CSV Guardado: {}",
        "files_added": "Añadidos {} nuevos archivos.",
        "files_cleared": "Selección de archivos limpiada.",
        "output_path_set": "Ruta de salida: {}",
        "fatal_error": "El proceso encontró un error fatal",
        "quota_error": "Se ha excedido el límite de la API diario (Código 429).",
        "api_error": "La clave API ingresada no es correcta (Código 400). Por favor, verifíquela e inténtelo de nuevo."
    }
}

class PDFToXLSXGUI:
    def __init__(self, root):
        self.root = root
        self.root.title(f"PDF Table Extractor v{VERSION}")
        self.root.geometry("800x800")
        self.root.minsize(750, 750)
        self.root.configure(bg="#FFFFFF")

        # Set style for modern look
        self.style = ttk.Style()
        self.style.theme_use('clam')
        self._configure_styles()

        self.pdf_files = []
        self.api_key = tk.StringVar()
        
        # Default output directory: extracted_tables in the current working directory
        default_out = os.path.join(os.getcwd(), "extracted_tables")
        self.output_dir = tk.StringVar(value=default_out)
        
        # Filename options
        self.excel_name = tk.StringVar(value="extracted_tables.xlsx")
        self.csv_name = tk.StringVar(value="extracted_tables.csv")
        self.md_name = tk.StringVar(value="extracted_tables.md")
        
        # Options
        self.save_excel = tk.BooleanVar(value=True)
        self.save_md = tk.BooleanVar(value=False)
        self.save_csv = tk.BooleanVar(value=False)
        self.clean_data = tk.BooleanVar(value=True)
        self._has_error = False  # Flag to track if errors occurred during processing
        
        # New State Variables
        self.lang = "ES" # Default to Spanish based on user request context
        self.current_prompt = DEFAULT_PROMPT.strip()
        self.ui_elements = {} # To hold references to widgets for language updates

        # Load Icons
        self.icons = {}
        self._load_icons()

        self._setup_ui()
        self._load_existing_api_key()

    def _configure_styles(self):
        # Configure the global background to white
        self.style.configure(".", background="#FFFFFF", foreground="#333333", font=("Segoe UI", 10))
        self.style.configure("TFrame", background="#FFFFFF")
        self.style.configure("TLabelframe", background="#FFFFFF", foreground="#555555")
        self.style.configure("TLabelframe.Label", background="#FFFFFF", foreground="#0078D4", font=("Segoe UI", 10, "bold"))
        self.style.configure("TLabel", background="#FFFFFF")
        self.style.configure("TButton", padding=5)
        self.style.configure("TCheckbutton", background="#FFFFFF")
        
        # Action button style
        self.style.configure("Action.TButton", font=("Segoe UI", 11, "bold"), foreground="#FFFFFF", background="#0078D4")
        self.style.map("Action.TButton", background=[('active', '#005A9E')])

    def _load_icons(self):
        icon_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "icons")
        icon_map = {
            "pdf": "pdf_icon.png",
            "excel": "excel_icon.png",
            "csv": "csv_icon.png",
            "md": "markdown_icon.png"
        }
        for key, filename in icon_map.items():
            path = os.path.join(icon_dir, filename)
            if os.path.exists(path):
                img = Image.open(path).resize((24, 24), Image.Resampling.LANCZOS)
                self.icons[key] = ImageTk.PhotoImage(img)
            else:
                self.icons[key] = None

    def _setup_ui(self):
        main_container = ttk.Frame(self.root, padding=20)
        main_container.pack(fill="both", expand=True)

        # Header
        self.ui_elements["title"] = ttk.Label(main_container, text=TEXTS[self.lang]["title"], font=("Segoe UI", 18, "bold"), foreground="#0078D4")
        self.ui_elements["title"].pack(pady=(0, 10))

        # 1. CONFIGURATION & SETTINGS
        config_frame = ttk.LabelFrame(main_container, text=TEXTS[self.lang]["config_section"], padding=15)
        config_frame.pack(fill="x", pady=5)
        self.ui_elements["config_section"] = config_frame
        
        # API Key Row
        api_row = ttk.Frame(config_frame)
        api_row.pack(fill="x", expand=True)
        
        self.ui_elements["api_key_label"] = ttk.Label(api_row, text=TEXTS[self.lang]["api_key"])
        self.ui_elements["api_key_label"].pack(side="left", padx=5)
        self.api_entry = ttk.Entry(api_row, textvariable=self.api_key, show="*")
        self.api_entry.pack(side="left", fill="x", expand=True, padx=5)
        
        self.ui_elements["save_key_btn"] = ttk.Button(api_row, text=TEXTS[self.lang]["save_key"], command=self._save_api_key)
        self.ui_elements["save_key_btn"].pack(side="left", padx=5)
        
        # Help link
        self.ui_elements["get_key_link"] = ttk.Label(config_frame, text=TEXTS[self.lang]["get_key"], foreground="#0078D4", cursor="hand2", font=("Segoe UI", 9, "underline"))
        self.ui_elements["get_key_link"].pack(anchor="w", padx=5, pady=(5, 10))
        self.ui_elements["get_key_link"].bind("<Button-1>", lambda e: webbrowser.open("https://aistudio.google.com/api-keys"))

        # settings row (Prompt / Language / iLovePDF)
        settings_row = ttk.Frame(config_frame)
        settings_row.pack(fill="x")

        # Edit Prompt Button
        self.ui_elements["edit_prompt_btn"] = ttk.Button(settings_row, text=TEXTS[self.lang]["edit_prompt"], command=self._open_prompt_editor)
        self.ui_elements["edit_prompt_btn"].pack(side="left", padx=5)

        # Language Switch
        ttk.Label(settings_row, text=" | ").pack(side="left")
        self.ui_elements["lang_label"] = ttk.Label(settings_row, text=TEXTS[self.lang]["language"])
        self.ui_elements["lang_label"].pack(side="left", padx=2)
        
        self.lang_btn = ttk.Button(settings_row, text="ES / EN", width=8, command=self._toggle_language)
        self.lang_btn.pack(side="left", padx=5)

        # iLovePDF Link
        ttk.Label(settings_row, text=" | ").pack(side="left")
        self.ui_elements["ilovepdf_link"] = ttk.Label(settings_row, text=TEXTS[self.lang]["ilovepdf"], foreground="#E91E63", cursor="hand2", font=("Segoe UI", 9, "underline"))
        self.ui_elements["ilovepdf_link"].pack(side="left", padx=5)
        self.ui_elements["ilovepdf_link"].bind("<Button-1>", lambda e: webbrowser.open("https://www.ilovepdf.com/" + ("es" if self.lang == "ES" else "")))

        # 2. PDF SELECTION
        file_frame = ttk.LabelFrame(main_container, text=TEXTS[self.lang]["pdf_selection"], padding=15)
        file_frame.pack(fill="x", pady=5)
        self.ui_elements["pdf_selection"] = file_frame
        
        top_file = ttk.Frame(file_frame)
        top_file.pack(fill="x")
        if self.icons["pdf"]:
            tk.Label(top_file, image=self.icons["pdf"], bg="#FFFFFF").pack(side="left", padx=2)
        self.ui_elements["select_files_label"] = ttk.Label(top_file, text=TEXTS[self.lang]["select_files"])
        self.ui_elements["select_files_label"].pack(side="left", padx=5)
        
        # Buttons
        self.ui_elements["add_files_btn"] = ttk.Button(top_file, text=TEXTS[self.lang]["add_files"], command=self._add_files)
        self.ui_elements["add_files_btn"].pack(side="right", padx=5)
        self.ui_elements["clear_btn"] = ttk.Button(top_file, text=TEXTS[self.lang]["clear"], command=self._clear_files)
        self.ui_elements["clear_btn"].pack(side="right", padx=5)

        # Compact file display
        self.file_listbox = tk.Listbox(file_frame, height=3, bd=1, relief="solid", highlightthickness=0, font=("Segoe UI", 8), bg="#FDFDFD")
        self.file_listbox.pack(fill="x", pady=(10, 0))
        
        scrollbar = ttk.Scrollbar(self.file_listbox, orient="vertical", command=self.file_listbox.yview)
        self.file_listbox.config(yscrollcommand=scrollbar.set)

        # 3. Output Configuration
        output_frame = ttk.LabelFrame(main_container, text=TEXTS[self.lang]["output_config"], padding=15)
        output_frame.pack(fill="x", pady=5)
        self.ui_elements["output_config"] = output_frame
        
        # Directory selection
        dir_frame = ttk.Frame(output_frame)
        dir_frame.pack(fill="x", pady=5)
        self.ui_elements["output_folder_label"] = ttk.Label(dir_frame, text=TEXTS[self.lang]["output_folder"])
        self.ui_elements["output_folder_label"].pack(side="left", padx=5)
        ttk.Entry(dir_frame, textvariable=self.output_dir).pack(side="left", fill="x", expand=True, padx=5)
        self.ui_elements["browse_btn"] = ttk.Button(dir_frame, text=TEXTS[self.lang]["browse"], command=self._browse_output_dir)
        self.ui_elements["browse_btn"].pack(side="left")
        
        # Filenames
        name_container = ttk.Frame(output_frame)
        name_container.pack(fill="x", pady=10)

        # Excel field
        exc_f = ttk.Frame(name_container)
        exc_f.pack(fill="x", pady=2)
        if self.icons["excel"]: tk.Label(exc_f, image=self.icons["excel"], bg="#FFFFFF").pack(side="left", padx=5)
        self.ui_elements["excel_name_label"] = ttk.Label(exc_f, text=TEXTS[self.lang]["excel_name"], width=15)
        self.ui_elements["excel_name_label"].pack(side="left")
        ttk.Entry(exc_f, textvariable=self.excel_name).pack(side="left", fill="x", expand=True, padx=5)

        # CSV field
        csv_f = ttk.Frame(name_container)
        csv_f.pack(fill="x", pady=2)
        if self.icons["csv"]: tk.Label(csv_f, image=self.icons["csv"], bg="#FFFFFF").pack(side="left", padx=5)
        self.ui_elements["csv_name_label"] = ttk.Label(csv_f, text=TEXTS[self.lang]["csv_name"], width=15)
        self.ui_elements["csv_name_label"].pack(side="left")
        ttk.Entry(csv_f, textvariable=self.csv_name).pack(side="left", fill="x", expand=True, padx=5)

        # MD field
        md_f = ttk.Frame(name_container)
        md_f.pack(fill="x", pady=2)
        if self.icons["md"]: tk.Label(md_f, image=self.icons["md"], bg="#FFFFFF").pack(side="left", padx=5)
        self.ui_elements["md_name_label"] = ttk.Label(md_f, text=TEXTS[self.lang]["md_name"], width=15)
        self.ui_elements["md_name_label"].pack(side="left")
        ttk.Entry(md_f, textvariable=self.md_name).pack(side="left", fill="x", expand=True, padx=5)

        # 4. Options Section
        opt_frame = ttk.LabelFrame(main_container, text=TEXTS[self.lang]["options_section"], padding=15)
        opt_frame.pack(fill="x", pady=5)
        self.ui_elements["options_section"] = opt_frame
        
        self.ui_elements["opt_excel"] = ttk.Checkbutton(opt_frame, text=TEXTS[self.lang]["opt_excel"], variable=self.save_excel)
        self.ui_elements["opt_excel"].pack(side="left", padx=10)
        self.ui_elements["opt_md"] = ttk.Checkbutton(opt_frame, text=TEXTS[self.lang]["opt_md"], variable=self.save_md)
        self.ui_elements["opt_md"].pack(side="left", padx=10)
        self.ui_elements["opt_csv"] = ttk.Checkbutton(opt_frame, text=TEXTS[self.lang]["opt_csv"], variable=self.save_csv)
        self.ui_elements["opt_csv"].pack(side="left", padx=10)
        self.ui_elements["opt_normalize"] = ttk.Checkbutton(opt_frame, text=TEXTS[self.lang]["opt_normalize"], variable=self.clean_data)
        self.ui_elements["opt_normalize"].pack(side="left", padx=10)

        # 5. Action Section
        action_frame = ttk.Frame(main_container, padding=10)
        action_frame.pack(fill="x")
        
        self.start_btn = ttk.Button(action_frame, text=TEXTS[self.lang]["start_btn"], style="Action.TButton", command=self._start_processing)
        self.ui_elements["start_btn"] = self.start_btn
        self.start_btn.pack(side="top", fill="x", pady=5)
        
        self.progress = ttk.Progressbar(action_frame, orient="horizontal", mode="determinate")
        self.progress.pack(fill="x", pady=10)

        # 6. Log Console
        log_frame = ttk.LabelFrame(main_container, text=TEXTS[self.lang]["status_log"], padding=10)
        log_frame.pack(fill="both", expand=True, pady=10)
        self.ui_elements["status_log"] = log_frame
        
        self.log_area = scrolledtext.ScrolledText(log_frame, height=5, font=("Consolas", 9), bg="#F9F9F9", fg="#333333", bd=0)
        self.log_area.pack(fill="both", expand=True)
        self.log_area.config(state="disabled")

    def _toggle_language(self):
        self.lang = "EN" if self.lang == "ES" else "ES"
        self._update_ui_language()
        self._log(f"Language changed to: {self.lang}")

    def _update_ui_language(self):
        t = TEXTS[self.lang]
        # Direct widgets
        self.ui_elements["title"].config(text=t["title"])
        self.ui_elements["config_section"].config(text=t["config_section"])
        self.ui_elements["api_key_label"].config(text=t["api_key"])
        self.ui_elements["save_key_btn"].config(text=t["save_key"])
        self.ui_elements["get_key_link"].config(text=t["get_key"])
        self.ui_elements["edit_prompt_btn"].config(text=t["edit_prompt"])
        self.ui_elements["lang_label"].config(text=t["language"])
        self.ui_elements["ilovepdf_link"].config(text=t["ilovepdf"])
        self.ui_elements["pdf_selection"].config(text=t["pdf_selection"])
        self.ui_elements["select_files_label"].config(text=t["select_files"])
        self.ui_elements["add_files_btn"].config(text=t["add_files"])
        self.ui_elements["clear_btn"].config(text=t["clear"])
        self.ui_elements["output_config"].config(text=t["output_config"])
        self.ui_elements["output_folder_label"].config(text=t["output_folder"])
        self.ui_elements["browse_btn"].config(text=t["browse"])
        self.ui_elements["excel_name_label"].config(text=t["excel_name"])
        self.ui_elements["csv_name_label"].config(text=t["csv_name"])
        self.ui_elements["md_name_label"].config(text=t["md_name"])
        self.ui_elements["options_section"].config(text=t["options_section"])
        self.ui_elements["opt_excel"].config(text=t["opt_excel"])
        self.ui_elements["opt_md"].config(text=t["opt_md"])
        self.ui_elements["opt_csv"].config(text=t["opt_csv"])
        self.ui_elements["opt_normalize"].config(text=t["opt_normalize"])
        self.ui_elements["start_btn"].config(text=t["start_btn"])
        self.ui_elements["status_log"].config(text=t["status_log"])
        
        # Explicit update iLovePDF link based on lang
        self.ui_elements["ilovepdf_link"].unbind("<Button-1>")
        self.ui_elements["ilovepdf_link"].bind("<Button-1>", lambda e: webbrowser.open("https://www.ilovepdf.com/" + ("es" if self.lang == "ES" else "")))

    def _open_prompt_editor(self):
        editor = tk.Toplevel(self.root)
        editor.title(TEXTS[self.lang]["prompt_editor_title"])
        editor.geometry("600x400")
        editor.transient(self.root)
        editor.grab_set()

        content_f = ttk.Frame(editor, padding=10)
        content_f.pack(fill="both", expand=True)

        txt_area = scrolledtext.ScrolledText(content_f, font=("Segoe UI", 10), wrap="word")
        txt_area.pack(fill="both", expand=True, pady=(0, 10))
        txt_area.insert(tk.END, self.current_prompt)

        btn_f = ttk.Frame(content_f)
        btn_f.pack(fill="x")

        def save_prompt():
            self.current_prompt = txt_area.get("1.0", tk.END).strip()
            editor.destroy()
            self._log("Prompt updated.")

        def reset_prompt():
            txt_area.delete("1.0", tk.END)
            txt_area.insert(tk.END, DEFAULT_PROMPT.strip())

        ttk.Button(btn_f, text=TEXTS[self.lang]["save"], command=save_prompt).pack(side="right", padx=5)
        ttk.Button(btn_f, text=TEXTS[self.lang]["cancel"], command=editor.destroy).pack(side="right", padx=5)
        ttk.Button(btn_f, text=TEXTS[self.lang]["reset"], command=reset_prompt).pack(side="left", padx=5)

    def _log(self, message):
        self.log_area.config(state="normal")
        self.log_area.insert(tk.END, f"[{time.strftime('%H:%M:%S')}] {message}\n")
        self.log_area.see(tk.END)
        self.log_area.config(state="disabled")
        self.root.update_idletasks()

    def _load_existing_api_key(self):
        env_path = os.path.join(os.getcwd(), "api_key.env")
        if os.path.exists(env_path):
            load_dotenv(env_path)
            key = os.getenv("API_KEY")
            if key:
                self.api_key.set(key)
                msg = "Ready: API key loaded." if self.lang == "EN" else "Listo: Clave API cargada."
                self._log(msg)

    def _save_api_key(self):
        key = self.api_key.get().strip()
        if not key:
            messagebox.showwarning(TEXTS[self.lang]["warning"], TEXTS[self.lang]["no_key"])
            return
        
        env_path = os.path.join(os.getcwd(), "api_key.env")
        if not os.path.exists(env_path):
            with open(env_path, "w") as f:
                f.write(f"API_KEY={key}\n")
        else:
            set_key(env_path, "API_KEY", key)
        
        messagebox.showinfo(TEXTS[self.lang]["success"], TEXTS[self.lang]["key_saved"])
        self._log(TEXTS[self.lang]["key_saved"])

    def _add_files(self):
        files = filedialog.askopenfilenames(filetypes=[("PDF Files", "*.pdf")])
        if files:
            for f in files:
                full_path = os.path.abspath(f)
                if full_path not in self.pdf_files:
                    self.pdf_files.append(full_path)
                    self.file_listbox.insert(tk.END, os.path.basename(f))
            self._log(TEXTS[self.lang]["files_added"].format(len(files)))

    def _clear_files(self):
        self.pdf_files = []
        self.file_listbox.delete(0, tk.END)
        self._log(TEXTS[self.lang]["files_cleared"])

    def _browse_output_dir(self):
        directory = filedialog.askdirectory()
        if directory:
            self.output_dir.set(os.path.abspath(directory))
            self._log(TEXTS[self.lang]["output_path_set"].format(directory))

    def _start_processing(self):
        if not self.api_key.get().strip():
            messagebox.showerror(TEXTS[self.lang]["error"], TEXTS[self.lang]["no_key"])
            return
        
        # Check API Key length
        api_key_val = self.api_key.get().strip()
        if len(api_key_val) != 39:
            messagebox.showwarning(TEXTS[self.lang]["error"], TEXTS[self.lang]["key_length"])
            return

        if not self.pdf_files:
            messagebox.showerror(TEXTS[self.lang]["error"], TEXTS[self.lang]["no_files"])
            return
        
        out_dir = self.output_dir.get().strip()
        # Default dir creation if it doesn't exist
        if not os.path.isdir(out_dir):
            try:
                os.makedirs(out_dir, exist_ok=True)
                self._log(f"Created output directory: {os.path.basename(out_dir)}")
            except Exception as e:
                messagebox.showerror("Error", f"Could not create output directory: {e}")
                return

        self.start_btn.config(state="disabled")
        self.progress["value"] = 0
        self.progress["maximum"] = len(self.pdf_files)
        
        thread = threading.Thread(target=self._process_logic)
        thread.daemon = True
        thread.start()

    def _process_logic(self):
        key = self.api_key.get().strip()
        out_dir = self.output_dir.get().strip()
        self._has_error = False # Reset error flag
        
        try:
            client = genai.Client(api_key=key)
            for logger_name in ["google", "google.genai", "urllib3"]:
                logging.getLogger(logger_name).setLevel(logging.WARNING)
        except Exception as e:
            self._log(f"FAIL: Gemini initialization failed: {e}")
            self.root.after(0, lambda: self.start_btn.config(state="normal"))
            return

        try:
            excel_filename = self.excel_name.get().strip()
            if not excel_filename.endswith('.xlsx'): excel_filename += '.xlsx'
            excel_path = os.path.join(out_dir, excel_filename)
            
            writer = None
            if self.save_excel.get():
                writer = pd.ExcelWriter(excel_path, engine='openpyxl')
                pd.DataFrame([["Tables extracted from GUI Application"]]).to_excel(writer, sheet_name="Summary", index=False, header=False)

            for i, pdf_path in enumerate(self.pdf_files):
                file_name = os.path.basename(pdf_path)
                self._log(TEXTS[self.lang]["working_on"].format(file_name))
                
                all_results = []
                try:
                    with pdfplumber.open(pdf_path) as pdf:
                        for page in pdf.pages:
                            page_res = self._extract_from_page(client, page)
                            all_results.extend(page_res)
                    
                    if all_results:
                        processed = []
                        for res in all_results:
                            df = res['df']
                            if self.clean_data.get():
                                df = self._normalize_df(df)
                            processed.append({"df": df, "md": res['md']})
                        
                        all_dfs = [p['df'] for p in processed]
                        combined_df = pd.concat(all_dfs, ignore_index=True)
                        
                        short_name = os.path.splitext(file_name)[0]
                        
                        if writer:
                            sheet_label = short_name[:31].strip()
                            combined_df.to_excel(writer, sheet_name=sheet_label, index=False, header=False)
                        
                        if self.save_md.get():
                            md_filename = self.md_name.get().strip()
                            if not md_filename.endswith('.md'): md_filename += '.md'
                            md_base = f"{short_name}_{md_filename}" if len(self.pdf_files) > 1 else md_filename
                            md_path = os.path.join(out_dir, md_base)
                            with open(md_path, 'w', encoding='utf-8') as f:
                                f.write(f"# Extracted Tables for {file_name}\n\n")
                                for idx, p in enumerate(processed):
                                    f.write(f"## Table {idx+1}\n\n{p['md']}\n\n")
                            self._log(TEXTS[self.lang]["saved_md"].format(md_base))
                        
                        if self.save_csv.get():
                            csv_filename = self.csv_name.get().strip()
                            if not csv_filename.endswith('.csv'): csv_filename += '.csv'
                            csv_base = f"{short_name}_{csv_filename}" if len(self.pdf_files) > 1 else csv_filename
                            csv_path = os.path.join(out_dir, csv_base)
                            combined_df.to_csv(csv_path, index=False, header=False)
                            self._log(TEXTS[self.lang]["saved_csv"].format(csv_base))
                            
                        self._log(TEXTS[self.lang]["done"].format(file_name))
                    else:
                        self._log(TEXTS[self.lang]["skip"].format(file_name))
                
                except Exception as e:
                    self._log(f"ERROR: {file_name} -> {e}")
                    if "400" in str(e) or "429" in str(e):
                        raise e # Re-raise to show the error window and stop completely
                
                self.root.after(0, self._update_progress, i + 1)

            if writer:
                writer.close()
                if not self._has_error:
                    self._log(f"Results consolidated in EXCEL: {excel_filename}")

            if not self._has_error:
                self._log(TEXTS[self.lang]["all_tasks_done"])
                self.root.after(0, lambda: messagebox.showinfo(TEXTS[self.lang]["process_finished"], TEXTS[self.lang]["process_success"]))
            else:
                self._log(TEXTS[self.lang]["process_error"])
            
        except Exception as e:
            self._log(f"CRITICAL ERROR: {e}")
            if "400" in str(e):
                messagebox.showerror(TEXTS[self.lang]["error"], TEXTS[self.lang]["api_error"])
            elif "429" in str(e):
                messagebox.showerror(TEXTS[self.lang]["error"], TEXTS[self.lang]["quota_error"])
            else:
                messagebox.showerror(TEXTS[self.lang]["error"], f"{TEXTS[self.lang]['fatal_error']}: {e}")
        
        self.root.after(0, lambda: self.start_btn.config(state="normal"))

    def _update_progress(self, val):
        self.progress["value"] = val

    def _extract_from_page(self, client, page):
        self._log(TEXTS[self.lang]["analyzing_page"].format(page.page_number))
        try:
            img = page.to_image(resolution=300).original
            prompt = self.current_prompt
            
            max_retries = 1
            md_text = ""
            for attempt in range(max_retries):
                try:
                    response = client.models.generate_content(
                        model='gemini-3-flash-preview',
                        contents=[prompt, img]
                    )
                    if response and hasattr(response, 'text') and response.text:
                        md_text = response.text
                    break
                except Exception as e:
                    if "400" in str(e):
                        self._log(f"      ! Incorrect API Key (Code 400) / API Key incorrecta (Código 400).")
                        self._has_error = True
                        raise e # Re-raise to stop the loop for this file or process
                    elif "429" in str(e):
                        if attempt < max_retries - 1:
                            time.sleep((attempt + 1) * 10)
                        else:
                            self._log(f"      ! Daily quota exceeded (Code 429) / Límite diario excedido (Código 429).")
                            self._has_error = True
                            raise e
                    else:
                        raise e
            
            clean_md = md_text.replace("```markdown", "").replace("```", "").strip()
            if clean_md:
                df = self._parse_md(clean_md)
                if not df.empty:
                    return [{"df": df, "md": clean_md}]
        except Exception as e:
            if "400" in str(e) or "429" in str(e):
                # Propagate fatal API errors
                raise e
            self._log(f"      ! Page {page.page_number} error: {e}")
        return []

    def _parse_md(self, md_text):
        lines = md_text.strip().split('\n')
        data = []
        for line in lines:
            if '|' in line:
                cells = [c.strip() for c in line.split('|')]
                if cells and cells[0] == '': cells = cells[1:]
                if cells and cells[-1] == '': cells = cells[:-1]
                if all(set(c) <= {'-', ':', ' '} for c in cells): continue
                if cells: data.append(cells)
            else:
                if line.strip(): data.append([line.strip()])
        if not data: return pd.DataFrame()
        max_cols = max(len(row) for row in data)
        norm_data = [row + [''] * (max_cols - len(row)) for row in data]
        return pd.DataFrame(norm_data)

    def _normalize_df(self, df):
        def clean_cell(val):
            if pd.isna(val) or val is None: return ""
            s = str(val).strip()
            s = s.replace('$', '').replace('€', '').replace(',', '')
            try:
                if '.' in s: return float(s)
                return int(s)
            except ValueError:
                return s
        return df.apply(lambda col: col.map(clean_cell))

if __name__ == "__main__":
    root = tk.Tk()
    # Explicitly set the icon for the window if pdf icon exists
    try:
        icon_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "icons", "pdf_icon.png")
        if os.path.exists(icon_path):
            img = Image.open(icon_path)
            photo = ImageTk.PhotoImage(img)
            root.wm_iconphoto(True, photo)
    except:
        pass
        
    app = PDFToXLSXGUI(root)
    root.mainloop()
