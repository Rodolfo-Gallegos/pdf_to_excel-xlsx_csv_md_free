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
VERSION = "1.3.0"

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

        # Header (Updated as per user preference)
        header_label = ttk.Label(main_container, text="PDF to EXCEL/CSV/MD AI Extractor", font=("Segoe UI", 18, "bold"), foreground="#0078D4")
        header_label.pack(pady=(0, 20))

        # 1. API Key Section
        api_frame = ttk.LabelFrame(main_container, text=" 1. CONFIGURATION ", padding=15)
        api_frame.pack(fill="x", pady=5)
        
        api_top = ttk.Frame(api_frame)
        api_top.pack(fill="x", expand=True)
        
        ttk.Label(api_top, text="API Key:").pack(side="left", padx=5)
        self.api_entry = ttk.Entry(api_top, textvariable=self.api_key, show="*")
        self.api_entry.pack(side="left", fill="x", expand=True, padx=5)
        
        ttk.Button(api_top, text="Save Key", command=self._save_api_key).pack(side="left", padx=5)
        
        api_bottom = ttk.Frame(api_frame)
        api_bottom.pack(fill="x", pady=(8, 0))
        
        link_label = ttk.Label(api_bottom, text="➜ Get your free API key at Google AI Studio", foreground="#0078D4", cursor="hand2", font=("Segoe UI", 9, "underline"))
        link_label.pack(side="left", padx=5)
        link_label.bind("<Button-1>", lambda e: webbrowser.open("https://aistudio.google.com/api-keys"))

        # 2. File Selection Section (Refactored to be more compact)
        file_frame = ttk.LabelFrame(main_container, text=" 2. PDF SELECTION ", padding=15)
        file_frame.pack(fill="x", pady=5)
        
        top_file = ttk.Frame(file_frame)
        top_file.pack(fill="x")
        if self.icons["pdf"]:
            tk.Label(top_file, image=self.icons["pdf"], bg="#FFFFFF").pack(side="left", padx=2)
        ttk.Label(top_file, text="Select files to process:").pack(side="left", padx=5)
        
        # Multi-select button
        ttk.Button(top_file, text="+ Add Files", command=self._add_files).pack(side="right", padx=5)
        ttk.Button(top_file, text="🗑 Clear", command=self._clear_files).pack(side="right", padx=5)

        # Compact file display (Listbox with small height)
        self.file_listbox = tk.Listbox(file_frame, height=3, bd=1, relief="solid", highlightthickness=0, font=("Segoe UI", 8), bg="#FDFDFD")
        self.file_listbox.pack(fill="x", pady=(10, 0))
        
        scrollbar = ttk.Scrollbar(self.file_listbox, orient="vertical", command=self.file_listbox.yview)
        # Scrollbar only visible if needed, but for 3 lines we keep it simple
        self.file_listbox.config(yscrollcommand=scrollbar.set)

        # 3. Output Configuration
        output_frame = ttk.LabelFrame(main_container, text=" 3. OUTPUT CONFIGURATION ", padding=15)
        output_frame.pack(fill="x", pady=5)
        
        # Directory selection
        dir_frame = ttk.Frame(output_frame)
        dir_frame.pack(fill="x", pady=5)
        ttk.Label(dir_frame, text="Output Folder:").pack(side="left", padx=5)
        ttk.Entry(dir_frame, textvariable=self.output_dir).pack(side="left", fill="x", expand=True, padx=5)
        ttk.Button(dir_frame, text="Browse...", command=self._browse_output_dir).pack(side="left")
        
        # Filenames with Icons
        name_container = ttk.Frame(output_frame)
        name_container.pack(fill="x", pady=10)

        # Excel field
        exc_f = ttk.Frame(name_container)
        exc_f.pack(fill="x", pady=2)
        if self.icons["excel"]: tk.Label(exc_f, image=self.icons["excel"], bg="#FFFFFF").pack(side="left", padx=5)
        ttk.Label(exc_f, text="Excel Name:", width=12).pack(side="left")
        ttk.Entry(exc_f, textvariable=self.excel_name).pack(side="left", fill="x", expand=True, padx=5)

        # CSV field
        csv_f = ttk.Frame(name_container)
        csv_f.pack(fill="x", pady=2)
        if self.icons["csv"]: tk.Label(csv_f, image=self.icons["csv"], bg="#FFFFFF").pack(side="left", padx=5)
        ttk.Label(csv_f, text="CSV Name:", width=12).pack(side="left")
        ttk.Entry(csv_f, textvariable=self.csv_name).pack(side="left", fill="x", expand=True, padx=5)

        # MD field
        md_f = ttk.Frame(name_container)
        md_f.pack(fill="x", pady=2)
        if self.icons["md"]: tk.Label(md_f, image=self.icons["md"], bg="#FFFFFF").pack(side="left", padx=5)
        ttk.Label(md_f, text="MD Name:", width=12).pack(side="left")
        ttk.Entry(md_f, textvariable=self.md_name).pack(side="left", fill="x", expand=True, padx=5)

        # 4. Options Section
        opt_frame = ttk.LabelFrame(main_container, text=" 4. OPTIONS ", padding=15)
        opt_frame.pack(fill="x", pady=5)
        
        ttk.Checkbutton(opt_frame, text="Excel (.xlsx)", variable=self.save_excel).pack(side="left", padx=10)
        ttk.Checkbutton(opt_frame, text="Markdown (.md)", variable=self.save_md).pack(side="left", padx=10)
        ttk.Checkbutton(opt_frame, text="CSV (.csv)", variable=self.save_csv).pack(side="left", padx=10)
        ttk.Checkbutton(opt_frame, text="Normalize Data", variable=self.clean_data).pack(side="left", padx=10)

        # 5. Action Section
        action_frame = ttk.Frame(main_container, padding=10)
        action_frame.pack(fill="x")
        
        # Updated Button Text (Manual Change Preserved)
        self.start_btn = ttk.Button(action_frame, text=" START EXTRACTION ", style="Action.TButton", command=self._start_processing)
        self.start_btn.pack(side="top", fill="x", pady=5)
        
        self.progress = ttk.Progressbar(action_frame, orient="horizontal", mode="determinate")
        self.progress.pack(fill="x", pady=10)

        # 6. Log Console
        log_frame = ttk.LabelFrame(main_container, text=" STATUS LOG ", padding=10)
        log_frame.pack(fill="both", expand=True, pady=10)
        
        self.log_area = scrolledtext.ScrolledText(log_frame, height=5, font=("Consolas", 9), bg="#F9F9F9", fg="#333333", bd=0)
        self.log_area.pack(fill="both", expand=True)
        self.log_area.config(state="disabled")

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
                self._log("Ready: API key loaded.")

    def _save_api_key(self):
        key = self.api_key.get().strip()
        if not key:
            messagebox.showwarning("Warning", "Please enter an API Key first.")
            return
        
        env_path = os.path.join(os.getcwd(), "api_key.env")
        if not os.path.exists(env_path):
            with open(env_path, "w") as f:
                f.write(f"API_KEY={key}\n")
        else:
            set_key(env_path, "API_KEY", key)
        
        messagebox.showinfo("Success", "API Key saved to api_key.env")
        self._log("API key updated successfully.")

    def _add_files(self):
        files = filedialog.askopenfilenames(filetypes=[("PDF Files", "*.pdf")])
        if files:
            for f in files:
                full_path = os.path.abspath(f)
                if full_path not in self.pdf_files:
                    self.pdf_files.append(full_path)
                    self.file_listbox.insert(tk.END, os.path.basename(f))
            self._log(f"Added {len(files)} new files.")

    def _clear_files(self):
        self.pdf_files = []
        self.file_listbox.delete(0, tk.END)
        self._log("File selection cleared.")

    def _browse_output_dir(self):
        directory = filedialog.askdirectory()
        if directory:
            self.output_dir.set(os.path.abspath(directory))
            self._log(f"Output path: {directory}")

    def _start_processing(self):
        if not self.api_key.get().strip():
            messagebox.showerror("Error", "Gemini API Key is required.")
            return
        if not self.pdf_files:
            messagebox.showerror("Error", "Please add at least one PDF file.")
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
                self._log(f"Working on: {file_name}")
                
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
                            self._log(f"  + Saved MD: {md_base}")
                        
                        if self.save_csv.get():
                            csv_filename = self.csv_name.get().strip()
                            if not csv_filename.endswith('.csv'): csv_filename += '.csv'
                            csv_base = f"{short_name}_{csv_filename}" if len(self.pdf_files) > 1 else csv_filename
                            csv_path = os.path.join(out_dir, csv_base)
                            combined_df.to_csv(csv_path, index=False, header=False)
                            self._log(f"  + Saved CSV: {csv_base}")
                            
                        self._log(f"DONE: {file_name}")
                    else:
                        self._log(f"SKIP: No tables in {file_name}")
                
                except Exception as e:
                    self._log(f"ERROR: {file_name} -> {e}")
                
                self.root.after(0, self._update_progress, i + 1)

            if writer:
                writer.close()
                self._log(f"Results consolidated in EXCEL: {excel_filename}")

            self._log("SYSTEM: All tasks completed.")
            self.root.after(0, lambda: messagebox.showinfo("Process Finished", "The extraction process has completed successfully!"))
            
        except Exception as e:
            self._log(f"CRITICAL ERROR: {e}")
            messagebox.showerror("System Error", f"The process encountered a fatal error: {e}")
        
        self.root.after(0, lambda: self.start_btn.config(state="normal"))

    def _update_progress(self, val):
        self.progress["value"] = val

    def _extract_from_page(self, client, page):
        self._log(f"    - Analyzing page {page.page_number}...")
        try:
            img = page.to_image(resolution=300).original
            prompt = """
            Analyze this page and extract ALL tables you see.
            Even if the table looks like a screenshot or an embedded image, extract it.
            Return results strictly in Markdown format.
            Do not include any introductory text, titles outside the table, or comments.
            If no tables are found, return an empty string.
            """
            
            max_retries = 3
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
                    if "429" in str(e) and attempt < max_retries - 1:
                        time.sleep((attempt + 1) * 10)
                    else:
                        raise e
            
            clean_md = md_text.replace("```markdown", "").replace("```", "").strip()
            if clean_md:
                df = self._parse_md(clean_md)
                if not df.empty:
                    return [{"df": df, "md": clean_md}]
        except Exception as e:
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
