import os
import sys
import time
import random
import openpyxl
from copy import copy
from openpyxl.utils import get_column_letter
import requests
from dotenv import load_dotenv
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
import queue
import winsound

# Import core crawler logic from sharing module
import crawler_core

# Ensure UTF-8 output on Windows terminal (safeguarded against console-less execution environments)
if sys.stdout is not None:
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass

# Ensure Windows Taskbar displays the application icon instead of generic Python/Tkinter icon
if sys.platform == "win32":
    try:
        import ctypes
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("nextpay.crawler.studio.1.0")
    except Exception:
        pass

def get_resource_path(relative_path):
    """Get absolute path to resource, works for dev and for PyInstaller bundle."""
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(os.path.dirname(__file__))
        
    path = os.path.join(base_path, relative_path)
    if os.path.exists(path):
        return path
        
    exe_dir = os.path.abspath(os.path.dirname(sys.executable))
    path = os.path.join(exe_dir, relative_path)
    if os.path.exists(path):
        return path
        
    cwd_path = os.path.join(os.getcwd(), relative_path)
    if os.path.exists(cwd_path):
        return cwd_path
        
    return None

class CrawlerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Nextpay Crawler Studio")
        self.root.geometry("850x700")
        self.root.minsize(800, 650)
        self.root.configure(bg="#1e1e2e")
        
        # Paths configuration
        self.env_path = os.path.join(os.getcwd(), ".env")
        self.cache_file = os.path.join(os.getcwd(), "session_cache.json")
        
        # State variables
        self.is_running = False
        self.crawler_thread = None
        
        # Colors (Catppuccin Mocha style dark theme)
        self.bg_color = "#1e1e2e"
        self.card_color = "#252538"
        self.text_color = "white"
        self.text_muted = "#a6adc8"
        self.entry_bg = "#181825"
        self.entry_fg = "#ffffff"
        
        self.blue = "#89b4fa"
        self.blue_hover = "#b4befe"
        
        self.green = "#a6e3a1"
        self.green_hover = "#abe9b3"
        
        self.red = "#f38ba8"
        self.red_hover = "#f2cdcd"
        
        self.gray = "#313244"
        self.gray_hover = "#45475a"
        self.disabled_bg = "#45475a"
        self.disabled_fg = "#585b70"
        
        # Load Icon (with fallback to .ico or .png)
        self.load_app_icon()
                
        # Thread queues
        self.log_queue = queue.Queue()
        self.progress_queue = queue.Queue()
        self.status_queue = queue.Queue()
        
        # Setup GUI elements
        self.setup_styles()
        self.create_widgets()
        self.load_settings()
        
        # Start queue checking loop
        self.process_queues()
        
    def load_app_icon(self):
        icon_ico = get_resource_path("software-agent.ico")
        icon_png = get_resource_path("software-agent.png")
        
        try:
            if icon_ico and os.path.exists(icon_ico):
                self.root.iconbitmap(default=icon_ico)
            if icon_png and os.path.exists(icon_png):
                self.icon_img = tk.PhotoImage(file=icon_png)
                self.root.iconphoto(True, self.icon_img)
            elif icon_ico and os.path.exists(icon_ico):
                self.root.iconbitmap(icon_ico)
        except Exception as e:
            print(f"Could not load icon: {e}")
            
    def setup_styles(self):
        # Configure overall style
        self.style = ttk.Style()
        self.style.theme_use("clam")
        
        # Progressbar styling
        self.style.configure("Custom.Horizontal.TProgressbar",
                             troughcolor="#11111b",
                             background=self.blue,
                             thickness=15,
                             borderwidth=0)
                             
    def create_widgets(self):
        # Header Container
        header_frame = tk.Frame(self.root, bg=self.bg_color, pady=15, padx=20)
        header_frame.pack(fill="x")
        
        title_label = tk.Label(
            header_frame,
            text="NEXTPAY CRAWLER STUDIO",
            font=("Segoe UI", 18, "bold"),
            fg=self.blue,
            bg=self.bg_color
        )
        title_label.pack(anchor="w")
        
        subtitle_label = tk.Label(
            header_frame,
            text="Extract wholesaler dealer prices and inventory count directly using openpyxl.",
            font=("Segoe UI", 10),
            fg=self.text_muted,
            bg=self.bg_color
        )
        subtitle_label.pack(anchor="w", pady=(2, 0))
        
        # Main Scrollable Panel Frame
        main_frame = tk.Frame(self.root, bg=self.bg_color, padx=20, pady=5)
        main_frame.pack(fill="both", expand=True)
        
        # 1. Config Card Frame
        config_frame = tk.LabelFrame(
            main_frame,
            text=" File Settings ",
            font=("Segoe UI", 10, "bold"),
            fg=self.blue,
            bg=self.card_color,
            bd=1,
            relief="solid",
            padx=15,
            pady=15
        )
        config_frame.pack(fill="x", pady=(0, 15))
        
        # Input Path Selector Row
        input_label = tk.Label(
            config_frame,
            text="Input Excel File (contains product codes):",
            font=("Segoe UI", 9, "bold"),
            fg=self.text_color,
            bg=self.card_color
        )
        input_label.pack(anchor="w")
        
        input_row = tk.Frame(config_frame, bg=self.card_color)
        input_row.pack(fill="x", pady=(5, 12))
        
        self.input_path_var = tk.StringVar()
        self.input_entry = tk.Entry(
            input_row,
            textvariable=self.input_path_var,
            font=("Segoe UI", 9),
            bg=self.entry_bg,
            readonlybackground=self.entry_bg,
            fg=self.entry_fg,
            disabledforeground=self.entry_fg,
            insertbackground=self.entry_fg,
            relief="flat",
            bd=5,
            state="readonly"
        )
        self.input_entry.pack(side="left", fill="x", expand=True, padx=(0, 10))
        
        self.input_browse_btn = self.create_button(
            input_row,
            text="Browse...",
            bg=self.blue,
            hover_bg=self.blue_hover,
            fg="#11111b",
            command=self.browse_input_file
        )
        self.input_browse_btn.pack(side="right")
        
        # Output Path Selector Row
        output_label = tk.Label(
            config_frame,
            text="Output Excel File (results saved here):",
            font=("Segoe UI", 9, "bold"),
            fg=self.text_color,
            bg=self.card_color
        )
        output_label.pack(anchor="w")
        
        output_row = tk.Frame(config_frame, bg=self.card_color)
        output_row.pack(fill="x", pady=(5, 0))
        
        self.output_path_var = tk.StringVar()
        self.output_entry = tk.Entry(
            output_row,
            textvariable=self.output_path_var,
            font=("Segoe UI", 9),
            bg=self.entry_bg,
            readonlybackground=self.entry_bg,
            fg=self.entry_fg,
            disabledforeground=self.entry_fg,
            insertbackground=self.entry_fg,
            relief="flat",
            bd=5,
            state="readonly"
        )
        self.output_entry.pack(side="left", fill="x", expand=True, padx=(0, 10))
        
        self.output_browse_btn = self.create_button(
            output_row,
            text="Browse...",
            bg=self.blue,
            hover_bg=self.blue_hover,
            fg="#11111b",
            command=self.browse_output_file
        )
        self.output_browse_btn.pack(side="right")
        
        # 2. Control & Progress Frame
        control_frame = tk.Frame(main_frame, bg=self.card_color, bd=1, relief="solid", padx=15, pady=15)
        control_frame.pack(fill="x", pady=(0, 15))
        
        # Buttons Row
        btn_row = tk.Frame(control_frame, bg=self.card_color)
        btn_row.pack(fill="x", pady=(0, 12))
        
        self.start_btn = self.create_button(
            btn_row,
            text="▶ Start Crawler",
            bg=self.green,
            hover_bg=self.green_hover,
            fg="#11111b",
            command=self.start_crawling,
            font=("Segoe UI", 10, "bold")
        )
        self.start_btn.pack(side="left", padx=(0, 10))
        
        self.stop_btn = self.create_button(
            btn_row,
            text="⏹ Stop",
            bg=self.gray,
            hover_bg=self.gray_hover,
            fg=self.red,
            command=self.stop_crawling,
            font=("Segoe UI", 10, "bold")
        )
        self.set_btn_state(self.stop_btn, "disabled", self.gray, self.red)
        self.stop_btn.pack(side="left", padx=(0, 10))
        
        self.excel_btn = self.create_button(
            btn_row,
            text="📊 Open Output File",
            bg=self.gray,
            hover_bg=self.gray_hover,
            fg=self.text_color,
            command=self.show_in_excel,
            font=("Segoe UI", 10, "bold")
        )
        self.excel_btn.pack(side="right")
        
        # Progress Section
        self.progress_label = tk.Label(
            control_frame,
            text="Status: Ready",
            font=("Segoe UI", 9, "bold"),
            fg=self.text_color,
            bg=self.card_color
        )
        self.progress_label.pack(anchor="w", pady=(0, 5))
        
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(
            control_frame,
            variable=self.progress_var,
            maximum=100,
            style="Custom.Horizontal.TProgressbar"
        )
        self.progress_bar.pack(fill="x")
        
        # 3. Log Console Frame
        log_frame = tk.Frame(main_frame, bg=self.card_color, bd=1, relief="solid", padx=15, pady=15)
        log_frame.pack(fill="both", expand=True, pady=(0, 10))
        
        log_title = tk.Label(
            log_frame,
            text="Live Execution Logs",
            font=("Segoe UI", 9, "bold"),
            fg=self.blue,
            bg=self.card_color
        )
        log_title.pack(anchor="w", pady=(0, 5))
        
        # Scrolled Text Box
        self.log_area = tk.Text(
            log_frame,
            font=("Consolas", 9),
            bg="#11111b",
            fg=self.text_color,
            insertbackground=self.text_color,
            relief="flat",
            bd=5,
            state="disabled",
            wrap="word"
        )
        self.log_area.pack(side="left", fill="both", expand=True)
        
        scrollbar = tk.Scrollbar(log_frame, command=self.log_area.yview, bg=self.card_color)
        scrollbar.pack(side="right", fill="y")
        self.log_area.configure(yscrollcommand=scrollbar.set)
        
    def create_button(self, parent, text, bg, hover_bg, fg, command, font=("Segoe UI", 9, "bold")):
        btn = tk.Button(
            parent,
            text=text,
            bg=bg,
            fg=fg,
            command=command,
            font=font,
            relief="flat",
            bd=0,
            activebackground=hover_bg,
            activeforeground=fg,
            padx=15,
            pady=7,
            cursor="hand2"
        )
        btn.default_bg = bg
        btn.hover_bg = hover_bg
        btn.default_fg = fg
        
        def on_enter(e):
            if btn["state"] == "normal":
                btn.configure(bg=btn.hover_bg)
                
        def on_leave(e):
            if btn["state"] == "normal":
                btn.configure(bg=getattr(btn, 'default_bg', bg))
                
        btn.bind("<Enter>", on_enter)
        btn.bind("<Leave>", on_leave)
        return btn
        
    def set_btn_state(self, btn, state, normal_bg=None, normal_fg=None):
        if state == "disabled":
            btn.configure(state="disabled", bg=self.disabled_bg, fg=self.disabled_fg)
        else:
            eff_bg = normal_bg or getattr(btn, 'default_bg', self.gray)
            eff_fg = normal_fg or getattr(btn, 'default_fg', self.text_color)
            btn.default_bg = eff_bg
            btn.default_fg = eff_fg
            btn.configure(state="normal", bg=eff_bg, fg=eff_fg, activeforeground=eff_fg)
            
    def load_settings(self):
        if not os.path.exists(self.env_path):
            # Create a blank env if not exists
            with open(self.env_path, "w", encoding="utf-8") as f:
                f.write("# Portal credentials and configuration\n")
                
        load_dotenv(dotenv_path=self.env_path, override=True)
        
        input_path = os.getenv("INPUT_EXCEL_PATH", "input.xlsx")
        output_path = os.getenv("OUTPUT_EXCEL_PATH", "results.xlsx")
        
        # Absolute paths mapping if relative
        if input_path and not os.path.isabs(input_path):
            input_path = os.path.abspath(input_path)
        if output_path and not os.path.isabs(output_path):
            output_path = os.path.abspath(output_path)
            
        self.input_path_var.set(input_path)
        self.output_path_var.set(output_path)
        
        self.write_log(f"Settings loaded. Input file: {input_path}")
        self.write_log(f"Settings loaded. Output file: {output_path}")
        
    def save_settings(self):
        input_path = self.input_path_var.get()
        output_path = self.output_path_var.get()
        
        env_lines = []
        if os.path.exists(self.env_path):
            with open(self.env_path, "r", encoding="utf-8") as f:
                env_lines = f.readlines()
                
        new_lines = []
        input_written = False
        output_written = False
        
        for line in env_lines:
            stripped = line.strip()
            if stripped.startswith("INPUT_EXCEL_PATH="):
                new_lines.append(f"INPUT_EXCEL_PATH={input_path}\n")
                input_written = True
            elif stripped.startswith("OUTPUT_EXCEL_PATH="):
                new_lines.append(f"OUTPUT_EXCEL_PATH={output_path}\n")
                output_written = True
            else:
                new_lines.append(line)
                
        if not input_written:
            new_lines.append(f"INPUT_EXCEL_PATH={input_path}\n")
        if not output_written:
            new_lines.append(f"OUTPUT_EXCEL_PATH={output_path}\n")
            
        with open(self.env_path, "w", encoding="utf-8") as f:
            f.writelines(new_lines)
            
        self.write_log("Paths saved to .env configuration.")
        
    def browse_input_file(self):
        filename = filedialog.askopenfilename(
            title="Select Input Excel File",
            filetypes=[("Excel Files", "*.xlsx *.xls")]
        )
        if filename:
            normalized_path = os.path.abspath(filename)
            self.input_path_var.set(normalized_path)
            self.write_log(f"Input file selected: {normalized_path}")
            self.save_settings()
            
    def browse_output_file(self):
        filename = filedialog.asksaveasfilename(
            title="Select Output Excel File Location",
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx")]
        )
        if filename:
            normalized_path = os.path.abspath(filename)
            self.output_path_var.set(normalized_path)
            self.write_log(f"Output file selected: {normalized_path}")
            self.save_settings()
            
    def write_log(self, text):
        timestamp = time.strftime("%H:%M:%S")
        self.log_area.configure(state="normal")
        self.log_area.insert(tk.END, f"[{timestamp}] {text}\n")
        self.log_area.see(tk.END)
        self.log_area.configure(state="disabled")
        
    def process_queues(self):
        # 1. Logs Queue
        try:
            while True:
                msg = self.log_queue.get_nowait()
                self.write_log(msg)
        except queue.Empty:
            pass
            
        # 2. Progress Queue
        try:
            while True:
                current, total, prod_code = self.progress_queue.get_nowait()
                if total > 0:
                    percent = (current / total) * 100
                    self.progress_var.set(percent)
                    self.progress_label.configure(
                        text=f"Progress: {current} / {total} products ({percent:.1f}%) | Processing: {prod_code}"
                    )
                else:
                    self.progress_var.set(0)
                    self.progress_label.configure(text="Progress: 0 / 0 (0%)")
        except queue.Empty:
            pass
            
        # 3. Status Queue
        try:
            while True:
                status_type, detail = self.status_queue.get_nowait()
                if status_type == "SUCCESS":
                    self.play_sound(success=True)
                    messagebox.showinfo("Success", f"Process complete!\n{detail}")
                    self.on_process_ended()
                elif status_type == "ERROR":
                    self.play_sound(success=False)
                    messagebox.showerror("Error", f"An error occurred:\n{detail}")
                    self.on_process_ended()
                elif status_type == "STOPPED":
                    self.play_sound(success=True)
                    messagebox.showinfo("Stopped", f"Process stopped by user.\n{detail}")
                    self.on_process_ended()
        except queue.Empty:
            pass
            
        self.root.after(100, self.process_queues)
        
    def play_sound(self, success=True):
        try:
            if success:
                winsound.MessageBeep(winsound.MB_ICONASTERISK)
            else:
                winsound.MessageBeep(winsound.MB_ICONHAND)
        except Exception:
            pass
            
    def show_in_excel(self):
        out_path = self.output_path_var.get()
        if os.path.exists(out_path):
            try:
                os.startfile(out_path)
                self.write_log(f"Opening output file: {out_path}")
            except Exception as e:
                self.play_sound(success=False)
                messagebox.showerror("Error", f"Failed to open file: {e}")
        else:
            self.play_sound(success=False)
            messagebox.showwarning("File Not Found", "Output Excel file does not exist yet. Run crawler to generate it.")
            
    def start_crawling(self):
        if self.is_running:
            return
            
        # Verify inputs
        input_file = self.input_path_var.get()
        output_file = self.output_path_var.get()
        
        if not input_file or not os.path.exists(input_file):
            self.play_sound(success=False)
            messagebox.showerror("Error", f"Input file not found at: {input_file}")
            return
            
        # Lock GUI controls
        self.is_running = True
        self.set_btn_state(self.start_btn, "disabled", self.green)
        self.set_btn_state(self.stop_btn, "normal", self.gray, self.red)
        self.set_btn_state(self.input_browse_btn, "disabled", self.blue)
        self.set_btn_state(self.output_browse_btn, "disabled", self.blue)
        
        self.progress_var.set(0)
        self.progress_label.configure(text="Status: Starting...")
        
        # Run background thread
        self.crawler_thread = threading.Thread(target=self.run_crawler_task, daemon=True)
        self.crawler_thread.start()
        
    def stop_crawling(self):
        if not self.is_running:
            return
        self.is_running = False
        self.progress_label.configure(text="Status: Stopping gracefully...")
        self.log_queue.put("Requesting crawl job termination...")
        self.set_btn_state(self.stop_btn, "disabled", self.gray, self.red)
        
    def on_process_ended(self):
        self.is_running = False
        self.set_btn_state(self.start_btn, "normal", self.green)
        self.set_btn_state(self.stop_btn, "disabled", self.gray, self.red)
        self.set_btn_state(self.input_browse_btn, "normal", self.blue)
        self.set_btn_state(self.output_browse_btn, "normal", self.blue)
        self.progress_label.configure(text="Status: Idle")
        
    # --- Background Thread Logic (Refactored to openpyxl and crawler_core) ---
    
    def run_crawler_task(self):
        # Reload dotenv settings from .env file inside thread
        load_dotenv(dotenv_path=self.env_path, override=True)
        login_url = os.getenv("LOGIN_URL")
        username = os.getenv("USERNAME")
        password = os.getenv("PASSWORD")
        
        if not login_url or not username or not password:
            self.log_queue.put("Error: Missing credentials in environment variables (.env file).")
            self.status_queue.put(("ERROR", "Missing credentials (LOGIN_URL, USERNAME, PASSWORD) in .env file."))
            return
            
        input_file = self.input_path_var.get()
        output_file = self.output_path_var.get()
        
        self.log_queue.put(f"Loading input file: {input_file}")
        
        # Load workbook safely using openpyxl
        wb = self.load_excel_safely(input_file)
        if wb is None:
            self.status_queue.put(("ERROR", "Failed to load input Excel. Check if it's corrupted or locked."))
            return
            
        sheet = wb.active
        
        # Detect header columns (row 1)
        header_row = []
        for col_idx in range(1, sheet.max_column + 1):
            header_row.append(sheet.cell(row=1, column=col_idx).value)
            
        code_col_idx = None
        for idx, val in enumerate(header_row):
            if val and str(val).lower().replace(" ", "") in ["ürünkodu", "urunkodu", "code"]:
                code_col_idx = idx + 1
                break
                
        if not code_col_idx:
            cols_str = ", ".join(map(str, [v for v in header_row if v is not None]))
            self.log_queue.put(f"Error: Could not find 'Ürün Kodu' column. Available columns: {cols_str}")
            self.status_queue.put(("ERROR", f"Could not find 'Ürün Kodu' column. Available columns:\n{cols_str}"))
            return
            
        # Ensure worksheet displays grid lines
        sheet.sheet_view.showGridLines = True
        if hasattr(sheet.views, 'sheetView') and sheet.views.sheetView:
            for sv in sheet.views.sheetView:
                sv.showGridLines = True
                
        ref_header_cell = sheet.cell(row=1, column=code_col_idx)
        
        # Find or append result columns
        price_col_idx = None
        qty_col_idx = None
        
        for idx, val in enumerate(header_row):
            if val == 'Alış Fiyatı':
                price_col_idx = idx + 1
            elif val == 'Ürün Adedi':
                qty_col_idx = idx + 1
                
        if not price_col_idx:
            price_col_idx = sheet.max_column + 1
            sheet.cell(row=1, column=price_col_idx, value="Alış Fiyatı")
            header_row.append("Alış Fiyatı")
            
        if not qty_col_idx:
            qty_col_idx = sheet.max_column + 1
            sheet.cell(row=1, column=qty_col_idx, value="Ürün Adedi")
            header_row.append("Ürün Adedi")
            
        # Apply layout, borders, font, and widths to result column headers
        for c_idx, title in [(price_col_idx, "Alış Fiyatı"), (qty_col_idx, "Ürün Adedi")]:
            h_cell = sheet.cell(row=1, column=c_idx)
            h_cell.value = title
            if ref_header_cell.has_style:
                if ref_header_cell.font: h_cell.font = copy(ref_header_cell.font)
                if ref_header_cell.border: h_cell.border = copy(ref_header_cell.border)
                if ref_header_cell.fill: h_cell.fill = copy(ref_header_cell.fill)
                if ref_header_cell.alignment: h_cell.alignment = copy(ref_header_cell.alignment)
            col_letter = get_column_letter(c_idx)
            curr_width = sheet.column_dimensions[col_letter].width if col_letter in sheet.column_dimensions else None
            if not curr_width or curr_width < 14:
                sheet.column_dimensions[col_letter].width = 16.0
            
        session = requests.Session()
        
        # Define callback for logging directly to the GUI queue
        def core_log(msg):
            self.log_queue.put(msg)
            
        if not crawler_core.login(session, username, password, self.cache_file, core_log):
            self.status_queue.put(("ERROR", "Authentication failed. Please verify credentials in your .env file."))
            return
            
        # Find all data rows (from row 2 onwards)
        total_rows = sheet.max_row - 1
        if total_rows <= 0:
            self.log_queue.put("No data rows found in Excel sheet.")
            self.status_queue.put(("ERROR", "Input Excel has no data rows to process."))
            return
            
        self.log_queue.put(f"Loaded {total_rows} items from Excel. Starting scraping loop...")
        
        success_count = 0
        
        for row_num in range(2, sheet.max_row + 1):
            # Check cancellation flag
            if not self.is_running:
                break
                
            cell_val = sheet.cell(row=row_num, column=code_col_idx).value
            product_code = str(cell_val).strip() if cell_val is not None else ""
            
            if not product_code or product_code.lower() in ["nan", "none", ""]:
                self.progress_queue.put((row_num - 1, total_rows, "[Skipping Blank Line]"))
                continue
                
            self.progress_queue.put((row_num - 1, total_rows, product_code))
            self.log_queue.put(f"[{row_num - 1}/{total_rows}] Scraping product: {product_code}")
            
            price, stock = crawler_core.search_and_extract(session, product_code, core_log)
            
            # Write results back directly
            price_cell = sheet.cell(row=row_num, column=price_col_idx, value=price)
            qty_cell = sheet.cell(row=row_num, column=qty_col_idx, value=stock)
            
            # Preserve cell layout and grid borders from reference row
            ref_data_cell = sheet.cell(row=row_num, column=code_col_idx)
            if ref_data_cell.has_style:
                for target_cell in (price_cell, qty_cell):
                    if ref_data_cell.border:
                        target_cell.border = copy(ref_data_cell.border)
                    if ref_data_cell.font and not target_cell.font:
                        target_cell.font = copy(ref_data_cell.font)
                    if ref_data_cell.fill and (not target_cell.fill or not target_cell.fill.fill_type):
                        target_cell.fill = copy(ref_data_cell.fill)
            
            self.log_queue.put(f"  -> Found Price: {price} | Stock: {stock}")
            success_count += 1
            
            # Mimic human delay, checking is_running flag in 100ms intervals
            if row_num - 1 < total_rows:
                sleep_time = random.uniform(2.0, 5.0)
                self.log_queue.put(f"  -> Waiting {sleep_time:.2f} seconds before next item...")
                
                # Check is_running during sleep
                steps = int(sleep_time * 10)
                for _ in range(steps):
                    if not self.is_running:
                        break
                    time.sleep(0.1)
                    
        # Ensure all data rows preserve the table grid layout and borders even if skipped
        for r_num in range(2, sheet.max_row + 1):
            ref_r = sheet.cell(row=r_num, column=code_col_idx)
            if ref_r.has_style and ref_r.border and ref_r.border.left and ref_r.border.left.style:
                for c_idx in (price_col_idx, qty_col_idx):
                    target_c = sheet.cell(row=r_num, column=c_idx)
                    if not target_c.border or not target_c.border.left or not target_c.border.left.style:
                        target_c.border = copy(ref_r.border)
                    if ref_r.font and (not target_c.font or not target_c.font.name):
                        target_c.font = copy(ref_r.font)

        # Ensure grid lines remain visible in output Excel
        sheet.sheet_view.showGridLines = True
        if hasattr(sheet.views, 'sheetView') and sheet.views.sheetView:
            for sv in sheet.views.sheetView:
                sv.showGridLines = True
                
        # Save output safely
        save_success = self.save_excel_safely(wb, output_file)
        
        if not self.is_running:
            if save_success:
                self.status_queue.put(("STOPPED", f"Process stopped mid-way. Partial results (processed {success_count} rows) saved to:\n{output_file}"))
            else:
                self.status_queue.put(("ERROR", "Process stopped but failed to save output Excel file safely."))
        else:
            if save_success:
                self.status_queue.put(("SUCCESS", f"Finished processing all {total_rows} items.\nResults saved to:\n{output_file}"))
            else:
                self.status_queue.put(("ERROR", "Crawling completed successfully, but failed to save output Excel file safely."))
                
    # --- File Input/Output routines ---
            
    def load_excel_safely(self, file_path):
        try:
            return openpyxl.load_workbook(file_path)
        except PermissionError:
            self.log_queue.put(f"Warning: {file_path} is locked by Excel. Trying alternative read stream...")
            temp_path = file_path.replace(".xlsx", "_temp_read.xlsx")
            cmd = f'powershell -Command "Copy-Item \'{file_path}\' \'{temp_path}\' -Force"'
            os.system(cmd)
            try:
                wb = openpyxl.load_workbook(temp_path)
                return wb
            except Exception as e:
                self.log_queue.put(f"Error copying and reading input Excel: {e}")
                return None
            finally:
                if os.path.exists(temp_path):
                    os.remove(temp_path)
        except Exception as e:
            self.log_queue.put(f"Error reading Excel file: {e}")
            return None
            
    def save_excel_safely(self, wb, output_path):
        try:
            wb.save(output_path)
            self.log_queue.put(f"Successfully saved output to: {output_path}")
            return True
        except PermissionError:
            self.log_queue.put(f"Warning: Write permission denied on {output_path} (is the file open?). Attempting backup path...")
            alternative_path = output_path.replace(".xlsx", f"_backup_{int(time.time())}.xlsx")
            try:
                wb.save(alternative_path)
                self.log_queue.put(f"Saved instead to backup path: {alternative_path}")
                return True
            except Exception as e:
                self.log_queue.put(f"Critical: Failed to save backup path: {e}")
                return False
        except Exception as e:
            self.log_queue.put(f"Error saving output Excel: {e}")
            return False

def main():
    root = tk.Tk()
    app = CrawlerApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
