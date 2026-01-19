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

# Modular imports
from src import config
from src.config import VERSION, DEFAULT_PROMPT, TEXTS
from src.logic.processor import normalize_df, parse_md, extract_from_page, parse_page_query

class PDFToXLSXGUI:
    def __init__(self, root):
        self.root = root
        self.root.title(f"PDF Table Extractor v{VERSION}")
        self.root.geometry("800x750") # Reduced height for better fit
        self.root.minsize(750, 650)
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
        self._has_error = False
        
        # State Variables
        self.lang = "EN" 
        self.current_prompt = DEFAULT_PROMPT
        self.ui_elements = {}

        # Load Icons
        self.icons = {}
        self._load_icons()

        self._setup_ui()
        self._load_existing_api_key()

    def _configure_styles(self):
        self.style.configure(".", background="#FFFFFF", foreground="#333333", font=("Segoe UI", 10))
        self.style.configure("TFrame", background="#FFFFFF")
        self.style.configure("TLabelframe", background="#FFFFFF", foreground="#555555")
        self.style.configure("TLabelframe.Label", background="#FFFFFF", foreground="#0078D4", font=("Segoe UI", 10, "bold"))
        self.style.configure("TLabel", background="#FFFFFF")
        self.style.configure("TButton", padding=5)
        self.style.configure("TCheckbutton", background="#FFFFFF")
        
        self.style.configure("Action.TButton", font=("Segoe UI", 11, "bold"), foreground="#FFFFFF", background="#0078D4")
        self.style.map("Action.TButton", background=[('active', '#005A9E')])

    def _load_icons(self):
        # Base directory is PDF_to_XLSX/
        base_dir = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
        icon_dir = os.path.join(base_dir, "src", "assets", "icons")
        icon_map = {
            "pdf": "pdf_to_excel.png", "excel": "excel_icon.png",
            "csv": "csv_icon.png", "md": "markdown_icon.png"
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

        self.ui_elements["title"] = ttk.Label(main_container, text=TEXTS[self.lang]["title"], font=("Segoe UI", 18, "bold"), foreground="#0078D4")
        self.ui_elements["title"].pack(pady=(0, 10))

        # 1. CONFIGURATION
        config_frame = ttk.LabelFrame(main_container, text=TEXTS[self.lang]["config_section"], padding=15)
        config_frame.pack(fill="x", pady=2)
        self.ui_elements["config_section"] = config_frame
        
        api_row = ttk.Frame(config_frame)
        api_row.pack(fill="x", expand=True)
        
        self.ui_elements["api_key_label"] = ttk.Label(api_row, text=TEXTS[self.lang]["api_key"])
        self.ui_elements["api_key_label"].pack(side="left", padx=5)
        self.api_entry = ttk.Entry(api_row, textvariable=self.api_key, show="*")
        self.api_entry.pack(side="left", fill="x", expand=True, padx=5)
        
        self.ui_elements["paste_key_btn"] = ttk.Button(api_row, text=TEXTS[self.lang]["paste_key"], width=8, command=self._paste_api_key)
        self.ui_elements["paste_key_btn"].pack(side="left", padx=2)
        
        self.ui_elements["clear_key_btn"] = ttk.Button(api_row, text=TEXTS[self.lang]["clear_key"], width=8, command=self._clear_api_key)
        self.ui_elements["clear_key_btn"].pack(side="left", padx=2)

        self.ui_elements["save_key_btn"] = ttk.Button(api_row, text=TEXTS[self.lang]["save_key"], command=self._save_api_key)
        self.ui_elements["save_key_btn"].pack(side="left", padx=5)
        
        self.ui_elements["get_key_link"] = ttk.Label(config_frame, text=TEXTS[self.lang]["get_key"], foreground="#0078D4", cursor="hand2", font=("Segoe UI", 9, "underline"))
        self.ui_elements["get_key_link"].pack(anchor="w", padx=5, pady=(5, 5))
        self.ui_elements["get_key_link"].bind("<Button-1>", lambda e: webbrowser.open("https://aistudio.google.com/api-keys"))

        settings_row = ttk.Frame(config_frame)
        settings_row.pack(fill="x")

        self.ui_elements["edit_prompt_btn"] = ttk.Button(settings_row, text=TEXTS[self.lang]["edit_prompt"], command=self._open_prompt_editor)
        self.ui_elements["edit_prompt_btn"].pack(side="left", padx=5)

        ttk.Label(settings_row, text=" | ").pack(side="left")
        self.ui_elements["lang_label"] = ttk.Label(settings_row, text=TEXTS[self.lang]["language"])
        self.ui_elements["lang_label"].pack(side="left", padx=2)
        
        self.lang_btn = ttk.Button(settings_row, text="ES / EN", width=8, command=self._toggle_language)
        self.lang_btn.pack(side="left", padx=5)

        ttk.Label(settings_row, text=" | ").pack(side="left")
        self.ui_elements["ilovepdf_link"] = ttk.Label(settings_row, text=TEXTS[self.lang]["ilovepdf"], foreground="#E91E63", cursor="hand2", font=("Segoe UI", 9, "underline"))
        self.ui_elements["ilovepdf_link"].pack(side="left", padx=5)
        self.ui_elements["ilovepdf_link"].bind("<Button-1>", lambda e: webbrowser.open("https://www.ilovepdf.com/" + ("es" if self.lang == "ES" else "")))

        # 2. PDF SELECTION
        file_frame = ttk.LabelFrame(main_container, text=TEXTS[self.lang]["pdf_selection"], padding=10)
        file_frame.pack(fill="x", pady=2)
        self.ui_elements["pdf_selection"] = file_frame
        
        top_file = ttk.Frame(file_frame)
        top_file.pack(fill="x")
        if self.icons.get("pdf"):
            tk.Label(top_file, image=self.icons["pdf"], bg="#FFFFFF").pack(side="left", padx=2)
        self.ui_elements["select_files_label"] = ttk.Label(top_file, text=TEXTS[self.lang]["select_files"])
        self.ui_elements["select_files_label"].pack(side="left", padx=5)
        
        self.ui_elements["add_files_btn"] = ttk.Button(top_file, text=TEXTS[self.lang]["add_files"], command=self._add_files)
        self.ui_elements["add_files_btn"].pack(side="right", padx=5)
        self.ui_elements["clear_btn"] = ttk.Button(top_file, text=TEXTS[self.lang]["clear"], command=self._clear_files)
        self.ui_elements["clear_btn"].pack(side="right", padx=5)

        self.file_listbox = tk.Listbox(file_frame, height=3, bd=1, relief="solid", highlightthickness=0, font=("Segoe UI", 8), bg="#FDFDFD")
        self.file_listbox.pack(fill="x", pady=(5, 0))
        
        scrollbar = ttk.Scrollbar(self.file_listbox, orient="vertical", command=self.file_listbox.yview)
        self.file_listbox.config(yscrollcommand=scrollbar.set)

        # 3. Output Configuration
        output_frame = ttk.LabelFrame(main_container, text=TEXTS[self.lang]["output_config"], padding=10)
        output_frame.pack(fill="x", pady=2)
        self.ui_elements["output_config"] = output_frame
        
        dir_frame = ttk.Frame(output_frame)
        dir_frame.pack(fill="x", pady=2)
        self.ui_elements["output_folder_label"] = ttk.Label(dir_frame, text=TEXTS[self.lang]["output_folder"])
        self.ui_elements["output_folder_label"].pack(side="left", padx=5)
        ttk.Entry(dir_frame, textvariable=self.output_dir).pack(side="left", fill="x", expand=True, padx=5)
        self.ui_elements["browse_btn"] = ttk.Button(dir_frame, text=TEXTS[self.lang]["browse"], command=self._browse_output_dir)
        self.ui_elements["browse_btn"].pack(side="left")
        
        name_container = ttk.Frame(output_frame)
        name_container.pack(fill="x", pady=2)

        for key, var, label_key in [("excel", self.excel_name, "excel_name"), ("csv", self.csv_name, "csv_name"), ("md", self.md_name, "md_name")]:
            row = ttk.Frame(name_container)
            row.pack(fill="x", pady=1)
            if self.icons.get(key): tk.Label(row, image=self.icons[key], bg="#FFFFFF").pack(side="left", padx=5)
            self.ui_elements[f"{key}_name_label"] = ttk.Label(row, text=TEXTS[self.lang][label_key], width=15)
            self.ui_elements[f"{key}_name_label"].pack(side="left")
            ttk.Entry(row, textvariable=var).pack(side="left", fill="x", expand=True, padx=5)

        # 4. Options Section
        opt_frame = ttk.LabelFrame(main_container, text=TEXTS[self.lang]["options_section"], padding=10)
        opt_frame.pack(fill="x", pady=2)
        self.ui_elements["options_section"] = opt_frame
        
        for k, v in [("opt_excel", self.save_excel), ("opt_md", self.save_md), ("opt_csv", self.save_csv), ("opt_normalize", self.clean_data)]:
            self.ui_elements[k] = ttk.Checkbutton(opt_frame, text=TEXTS[self.lang][k], variable=v)
            self.ui_elements[k].pack(side="left", padx=10)

        # 5. Action Section
        action_frame = ttk.Frame(main_container, padding=5)
        action_frame.pack(fill="x")
        
        self.start_btn = ttk.Button(action_frame, text=TEXTS[self.lang]["start_btn"], style="Action.TButton", command=self._start_processing)
        self.ui_elements["start_btn"] = self.start_btn
        self.start_btn.pack(side="top", fill="x", pady=2)
        
        self.progress = ttk.Progressbar(action_frame, orient="horizontal", mode="determinate")
        self.progress.pack(fill="x", pady=5)

        # 6. Log Console
        log_frame = ttk.LabelFrame(main_container, text=TEXTS[self.lang]["status_log"], padding=5)
        log_frame.pack(fill="both", expand=True, pady=2)
        self.ui_elements["status_log"] = log_frame
        
        self.log_area = scrolledtext.ScrolledText(log_frame, height=4, font=("Consolas", 9), bg="#F9F9F9", fg="#333333", bd=0)
        self.log_area.pack(fill="both", expand=True)
        self.log_area.config(state="disabled")

    def _paste_api_key(self):
        try:
            self.api_key.set(self.root.clipboard_get())
        except:
            pass

    def _clear_api_key(self):
        self.api_key.set("")

    def _toggle_language(self):
        self.lang = "EN" if self.lang == "ES" else "ES"
        self._update_ui_language()
        self._log(f"Language changed to: {self.lang}")

    def _update_ui_language(self):
        t = TEXTS[self.lang]
        
        # Mappings of widget references in ui_elements to TEXTS keys
        mappings = {
            "title": "title",
            "config_section": "config_section",
            "api_key_label": "api_key",
            "paste_key_btn": "paste_key",
            "clear_key_btn": "clear_key",
            "save_key_btn": "save_key",
            "get_key_link": "get_key",
            "edit_prompt_btn": "edit_prompt",
            "lang_label": "language",
            "ilovepdf_link": "ilovepdf",
            "pdf_selection": "pdf_selection",
            "select_files_label": "select_files",
            "add_files_btn": "add_files",
            "clear_btn": "clear",
            "output_config": "output_config",
            "output_folder_label": "output_folder",
            "browse_btn": "browse",
            "excel_name_label": "excel_name",
            "csv_name_label": "csv_name",
            "md_name_label": "md_name",
            "options_section": "options_section",
            "opt_excel": "opt_excel",
            "opt_md": "opt_md",
            "opt_csv": "opt_csv",
            "opt_normalize": "opt_normalize",
            "start_btn": "start_btn",
            "status_log": "status_log"
        }

        for widget_key, text_key in mappings.items():
            if widget_key in self.ui_elements:
                try:
                    self.ui_elements[widget_key].config(text=t[text_key])
                except:
                    pass

        self.ui_elements["ilovepdf_link"].unbind("<Button-1>")
        self.ui_elements["ilovepdf_link"].bind("<Button-1>", lambda e: webbrowser.open("https://www.ilovepdf.com/" + ("es" if self.lang == "ES" else "")))

    def _open_prompt_editor(self):
        editor = tk.Toplevel(self.root)
        editor.title(TEXTS[self.lang]["prompt_editor_title"])
        editor.geometry("550x400") # Balanced size
        editor.transient(self.root)
        editor.grab_set()

        content_f = ttk.Frame(editor, padding=10)
        content_f.pack(fill="both", expand=True)

        # Pack button frame first at the bottom to ensure visibility
        btn_f = ttk.Frame(content_f)
        btn_f.pack(side="bottom", fill="x", pady=(10, 0))

        # Pack text area to take remaining space at the top
        txt_area = scrolledtext.ScrolledText(content_f, font=("Segoe UI", 12), wrap="word")
        txt_area.pack(side="top", fill="both", expand=True)
        txt_area.insert(tk.END, self.current_prompt)

        def save_prompt():
            self.current_prompt = txt_area.get("1.0", tk.END).strip()
            editor.destroy()
            self._log("Prompt updated.")

        def reset_prompt():
            txt_area.delete("1.0", tk.END)
            txt_area.insert(tk.END, DEFAULT_PROMPT)

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
        base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__))) # points to src/
        env_path = os.path.join(base_dir, "api_key.env")
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
        base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__))) # points to src/
        env_path = os.path.join(base_dir, "api_key.env")
        if not os.path.exists(env_path):
            with open(env_path, "w") as f: f.write(f"API_KEY={key}\n")
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
        key = self.api_key.get().strip()
        if not key:
            messagebox.showerror(TEXTS[self.lang]["error"], TEXTS[self.lang]["no_key"])
            return
        if len(key) != 39:
            messagebox.showwarning(TEXTS[self.lang]["error"], TEXTS[self.lang]["key_length"])
            return
        if not self.pdf_files:
            messagebox.showerror(TEXTS[self.lang]["error"], TEXTS[self.lang]["no_files"])
            return
        
        out_dir = self.output_dir.get().strip()
        if not os.path.isdir(out_dir):
            try:
                os.makedirs(out_dir, exist_ok=True)
            except Exception as e:
                messagebox.showerror("Error", f"Could not create output directory: {e}")
                return

        self.start_btn.config(state="disabled")
        self.progress["value"] = 0
        self.progress["maximum"] = len(self.pdf_files)
        threading.Thread(target=self._process_logic, daemon=True).start()

    def _process_logic(self):
        key = self.api_key.get().strip()
        out_dir = self.output_dir.get().strip()
        self._has_error = False
        
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
                if os.path.exists(excel_path):
                    writer = pd.ExcelWriter(excel_path, engine='openpyxl', mode='a', if_sheet_exists='replace')
                else:
                    writer = pd.ExcelWriter(excel_path, engine='openpyxl')
                    pd.DataFrame([["Tables extracted from GUI Application"]]).to_excel(writer, sheet_name="Summary", index=False, header=False)

            tracker = {"has_error": False}
            for i, pdf_path in enumerate(self.pdf_files):
                file_name = os.path.basename(pdf_path)
                self._log(TEXTS[self.lang]["working_on"].format(file_name))
                
                all_results = []
                try:
                    with pdfplumber.open(pdf_path) as pdf:
                        total_pages = len(pdf.pages)
                        all_basenames = [os.path.basename(f) for f in self.pdf_files]
                        pages_to_process = parse_page_query(self.current_prompt, total_pages, file_name, all_basenames)
                        
                        if len(pages_to_process) < total_pages:
                            self._log(f"Selective Mode: Processing {len(pages_to_process)} specific pages.")

                        for p_idx in pages_to_process:
                            page = pdf.pages[p_idx]
                            page_res = extract_from_page(client, page, self.current_prompt, 
                                                       lambda p: self._log(TEXTS[self.lang]["analyzing_page"].format(p)),
                                                       tracker)
                            all_results.extend(page_res)
                    
                    if all_results:
                        processed = []
                        for res in all_results:
                            df = res['df']
                            if self.clean_data.get(): df = normalize_df(df)
                            processed.append({"df": df, "md": res['md']})
                        
                        all_dfs = [p['df'] for p in processed]
                        combined_df = pd.concat(all_dfs, ignore_index=True)
                        short_name = os.path.splitext(file_name)[0]
                        
                        if writer:
                            combined_df.to_excel(writer, sheet_name=short_name[:31].strip(), index=False, header=False)
                        
                        if self.save_md.get():
                            md_filename = self.md_name.get().strip()
                            if not md_filename.endswith('.md'): md_filename += '.md'
                            md_base = f"{short_name}_{md_filename}" if len(self.pdf_files) > 1 else md_filename
                            with open(os.path.join(out_dir, md_base), 'w', encoding='utf-8') as f:
                                f.write(f"# Extracted Tables for {file_name}\n\n")
                                for idx, p in enumerate(processed): f.write(f"## Table {idx+1}\n\n{p['md']}\n\n")
                            self._log(TEXTS[self.lang]["saved_md"].format(md_base))
                        
                        if self.save_csv.get():
                            csv_filename = self.csv_name.get().strip()
                            if not csv_filename.endswith('.csv'): csv_filename += '.csv'
                            csv_base = f"{short_name}_{csv_filename}" if len(self.pdf_files) > 1 else csv_filename
                            combined_df.to_csv(os.path.join(out_dir, csv_base), index=False, header=False)
                            self._log(TEXTS[self.lang]["saved_csv"].format(csv_base))
                            
                        self._log(TEXTS[self.lang]["done"].format(file_name))
                    else:
                        self._log(TEXTS[self.lang]["skip"].format(file_name))
                
                except Exception as e:
                    self._log(f"ERROR: {file_name} -> {e}")
                    if tracker["has_error"]: raise e
                
                self.root.after(0, lambda v=i+1: self.progress.config(value=v))

            if writer:
                writer.close()
                if not tracker["has_error"]: self._log(f"Results consolidated in EXCEL: {excel_filename}")

            if not tracker["has_error"]:
                self._log(TEXTS[self.lang]["all_tasks_done"])
                self.root.after(0, lambda: messagebox.showinfo(TEXTS[self.lang]["process_finished"], TEXTS[self.lang]["process_success"]))
            else:
                self._log(TEXTS[self.lang]["process_error"])
            
        except Exception as e:
            self._log(f"CRITICAL ERROR: {e}")
            err_type = "api_error" if "400" in str(e) else "quota_error" if "429" in str(e) else "api_leaked" if "403" in str(e) else "unknown_error"
            if err_type:
                messagebox.showerror(TEXTS[self.lang]["error"], TEXTS[self.lang][err_type])
            else:
                messagebox.showerror(TEXTS[self.lang]["error"], f"{TEXTS[self.lang]['fatal_error']}: {e}")
        
        self.root.after(0, lambda: self.start_btn.config(state="normal"))
