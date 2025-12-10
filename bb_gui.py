"""
Blaze Buddy - GUI Module

Enterprise Edition
"""

import tkinter as tk
from tkinter import ttk, filedialog, scrolledtext, messagebox
import threading
import queue
import os
from datetime import datetime
from bb_config_manager import ConfigManager
from bb_utils import APP_TITLE, resource_path
from bb_logic import run_logic_pandas
import bb_ingest

# --- CORPORATE THEME PALETTE ---
THEME = {
    "primary": "#2E8B57",    # SeaGreen (The "Greenline" Brand)
    "secondary": "#E0E0E0",  # Light Gray background
    "accent": "#4682B4",     # SteelBlue for interactions
    "text_dark": "#333333",  # Soft black
    "text_light": "#FFFFFF", # White
    "bg_main": "#F5F5F5",    # Very light gray app background
    "font_main": "Segoe UI",
}

class GreenlineApp:
    def __init__(self, root):
        self.root = root
        self.root.title(APP_TITLE)
        self.root.geometry("900x950") # Slightly taller for the new look
        self.config = ConfigManager()
        
        # Load Icon
        try:
            icon = resource_path("icon.ico")
            if os.path.exists(icon): self.root.iconbitmap(icon)
        except (FileNotFoundError, OSError, tk.TclError):
            pass
        
        # --- APPLY NEW THEME ---
        self.apply_theme()
        
        # Data vars (Workflow remains identical)
        self.files = {k: tk.StringVar() for k in ['inventory', 'sales', 'po', 'aglc']}
        self.loc_files = {}
        for loc in self.config.get_locations(active_only=True):
            self.loc_files[loc['id']] = tk.StringVar()
            
        # Load saved paths
        saved_paths = self.config.get_paths()
        for k, v in saved_paths.items():
            if k in self.files: self.files[k].set(v)
            if k in self.loc_files: self.loc_files[k].set(v)
            
        self.report_days = tk.StringVar(value="30")
        self.log_queue = queue.Queue()
        
        self.create_widgets()
        self._process_log_queue()

    def apply_theme(self):
        """Injects the 'Enterprise' look and feel."""
        style = ttk.Style()
        style.theme_use('clam') # Good base for customization
        
        # Global Font
        default_font = (THEME['font_main'], 10)
        header_font = (THEME['font_main'], 22, "bold")
        sub_font = (THEME['font_main'], 10, "italic")
        
        # Configure Colors & Fonts
        style.configure(".", background=THEME['bg_main'], font=default_font, foreground=THEME['text_dark'])
        
        # Frames
        style.configure("TFrame", background=THEME['bg_main'])
        
        # Header Style (The "Corporate" Top Bar)
        style.configure("Brand.TFrame", background=THEME['primary'])
        style.configure("Brand.TLabel", background=THEME['primary'], foreground=THEME['text_light'], font=header_font)
        style.configure("BrandSub.TLabel", background=THEME['primary'], foreground="#D3D3D3", font=sub_font)
        
        # Buttons
        style.configure("TButton", padding=6, font=(THEME['font_main'], 10, 'bold'))
        style.map("TButton",
            background=[('active', THEME['accent']), ('!disabled', THEME['secondary'])],
            foreground=[('active', THEME['text_light']), ('!disabled', THEME['text_dark'])]
        )
        
        # The "Run" Button (Big & distinct)
        style.configure("Action.TButton", font=(THEME['font_main'], 12, 'bold'), background=THEME['primary'], foreground=THEME['text_light'])
        style.map("Action.TButton", background=[('active', '#1E5E3A')]) # Darker green on hover

        # LabelFrames (Groups)
        style.configure("TLabelframe", background=THEME['bg_main'], borderwidth=2, relief="groove")
        style.configure("TLabelframe.Label", background=THEME['bg_main'], foreground=THEME['primary'], font=(THEME['font_main'], 11, 'bold'))

    def create_widgets(self):
        # --- 1. THE BRAND HEADER ---
        # This solid color block replaces the old plain text header
        hdr = ttk.Frame(self.root, style="Brand.TFrame", padding="20")
        hdr.pack(fill=tk.X)
        
        # Logo/Title Area
        title_box = ttk.Frame(hdr, style="Brand.TFrame")
        title_box.pack(side=tk.LEFT)
        
        ttk.Label(title_box, text=APP_TITLE.upper(), style="Brand.TLabel").pack(anchor='w')
        ttk.Label(title_box, text="Inventory Intelligence Suite v2.0", style="BrandSub.TLabel").pack(anchor='w')
        
        # Utility Buttons (Top Right, lighter touch)
        btns = ttk.Frame(hdr, style="Brand.TFrame")
        btns.pack(side=tk.RIGHT)
        ttk.Button(btns, text="⚙️ Config", command=self.open_settings).pack(side=tk.RIGHT, padx=5)
        ttk.Button(btns, text="📥 Sync", command=self.run_ingest).pack(side=tk.RIGHT, padx=5)
        ttk.Button(btns, text="📂 Load", command=self.auto_load).pack(side=tk.RIGHT, padx=5)

        # --- 2. MAIN DASHBOARD ---
        main = ttk.Frame(self.root, padding="25")
        main.pack(fill=tk.BOTH, expand=True)
        
        # Section: Parameters
        fp = ttk.LabelFrame(main, text=" ANALYSIS PARAMETERS ", padding="15")
        fp.pack(fill=tk.X, pady=(0, 15))
        
        p_row = ttk.Frame(fp)
        p_row.pack(fill=tk.X)
        ttk.Label(p_row, text="Sales History Range (Days):").pack(side=tk.LEFT)
        ttk.Entry(p_row, textvariable=self.report_days, width=8).pack(side=tk.LEFT, padx=10)
        ttk.Label(p_row, text="(Standard: 30)", foreground="#888").pack(side=tk.LEFT)

        # Section: Source Data
        lfr = ttk.LabelFrame(main, text=" SOURCE DATA ", padding="15")
        lfr.pack(fill=tk.X, pady=(0, 15))
        self.make_row(lfr, "Inventory Export", self.files['inventory'], 'inventory')
        self.make_row(lfr, "Sales Report", self.files['sales'], 'sales')
        self.make_row(lfr, "Purchase Order", self.files['po'], 'po')
        self.make_row(lfr, "AGLC Manual", self.files['aglc'], 'aglc')

        # Section: Logistics (Dynamic)
        lft = ttk.LabelFrame(main, text=" LOGISTICS & TRANSFERS ", padding="15")
        lft.pack(fill=tk.X, pady=(0, 20))
        
        locs = self.config.get_locations(active_only=True)
        if locs:
            for loc in locs:
                self.make_row(lft, f"To {loc['name']}", self.loc_files[loc['id']], loc['id'])
        else:
             ttk.Label(lft, text="No active locations configured.", foreground="red").pack()

        # --- 3. EXECUTION ZONE ---
        # The button is now styled differently ("Action.TButton") to stand out
        self.btn_run = ttk.Button(main, text="GENERATE ORDER RECOMMENDATIONS 🚀", style="Action.TButton", command=self.start_run)
        self.btn_run.pack(fill=tk.X, ipady=12, pady=(0, 15))
        
        self.prog = ttk.Progressbar(main, mode='indeterminate')
        self.prog.pack(fill=tk.X, pady=(0,15))
        
        # Log Console
        lfl = ttk.LabelFrame(main, text=" SYSTEM LOG ", padding="10")
        lfl.pack(fill=tk.BOTH, expand=True)
        
        self.log_txt = scrolledtext.ScrolledText(lfl, height=8, state='disabled', font=("Consolas", 9))
        self.log_txt.pack(fill=tk.BOTH, expand=True)
        self.log_txt.configure(background="#F0F0F0", foreground="#333333")

    def make_row(self, p, lbl, var, key):
        f = ttk.Frame(p)
        f.pack(fill=tk.X, pady=4)
        ttk.Label(f, text=lbl, width=18, anchor="w").pack(side=tk.LEFT)
        e = ttk.Entry(f, textvariable=var)
        e.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        ttk.Button(f, text="Browse", width=8, command=lambda: self.browse(var, key)).pack(side=tk.RIGHT)

    def browse(self, var, key):
        path = filedialog.askopenfilename()
        if path:
            var.set(path)
            self.save_paths()

    def save_paths(self):
        paths = {k: v.get() for k, v in self.files.items()}
        paths.update({k: v.get() for k, v in self.loc_files.items()})
        self.config.update_paths(paths)

    def log(self, msg): self.log_queue.put(msg)
    
    def _process_log_queue(self):
        try:
            while True:
                msg = self.log_queue.get_nowait()
                self.log_txt.config(state='normal')
                ts = datetime.now().strftime('%H:%M:%S')
                self.log_txt.insert(tk.END, f"[{ts}] {msg}\n")
                self.log_txt.see(tk.END)
                self.log_txt.config(state='disabled')
        except queue.Empty:
            pass
        except Exception as e:
            print(f"Log processing error: {e}")
        self.root.after(100, self._process_log_queue)

    def auto_load(self):
        d = filedialog.askdirectory()
        if not d: return
        self.log(f"Scanning directory: {d}...")
        try:
            files = [os.path.join(d,f) for f in os.listdir(d) if f.lower().endswith(('.csv','.xlsx'))]
            found = 0
            for f in files:
                n = os.path.basename(f).lower()
                if "inventory-export" in n: self.files['inventory'].set(f); found+=1
                elif "product-sales" in n: self.files['sales'].set(f); found+=1
                elif "purchase-order" in n: self.files['po'].set(f); found+=1
                elif "manual" in n and "order" in n: self.files['aglc'].set(f); found+=1
                elif "transfer" in n:
                    for lid, lvar in self.loc_files.items():
                        if lid in n: lvar.set(f); found+=1
            self.save_paths()
            self.log(f"Auto-load complete. Mapped {found} files.")
        except Exception as e:
            self.log(f"Error during auto-load: {e}")

    def run_ingest(self):
        self.log("Initializing data synchronization...")
        threading.Thread(target=lambda: bb_ingest.update_memory_bank(self.log), daemon=True).start()

    def open_settings(self):
        # Reusing your existing logic but ensuring it spawns correctly
        top = tk.Toplevel(self.root)
        top.title("System Configuration")
        top.geometry("600x700")
        
        sets = self.config.get_settings()
        entries = {}
        
        def add_sec(title, key):
            lf = ttk.LabelFrame(top, text=f" {title} ", padding=10)
            lf.pack(fill=tk.X, padx=10, pady=5)
            for subk, val in sets.get(key, {}).items():
                if isinstance(val, (int, float)):
                    f = ttk.Frame(lf)
                    f.pack(fill=tk.X, pady=2)
                    ttk.Label(f, text=subk.replace('_', ' ').title(), width=25).pack(side=tk.LEFT)
                    e = ttk.Entry(f)
                    e.insert(0, val)
                    e.pack(side=tk.RIGHT, expand=True, fill=tk.X)
                    entries[f"{key}.{subk}"] = e
        
        add_sec("Cannabis Rules", "cannabis_logic")
        add_sec("Accessory Rules", "accessory_logic")
        
        # Save Button
        def save():
            try:
                new_s = sets.copy()
                for k, e in entries.items():
                    val = e.get()
                    if '.' in k:
                        sec, sub = k.split('.')
                        new_s[sec][sub] = float(val)
                self.config.update_settings(new_s)
                top.destroy()
                self.log("Configuration updated successfully.")
            except Exception as ex: messagebox.showerror("Config Error", str(ex))
            
        ttk.Button(top, text="SAVE CONFIGURATION", command=save).pack(pady=20, fill=tk.X, padx=20)

    def start_run(self):
        if not self.files['inventory'].get() or not self.files['sales'].get():
            messagebox.showerror("Input Error", "Inventory and Sales files are required for analysis.")
            return
        
        paths = {k: v.get() for k, v in self.files.items()}
        paths.update({k: v.get() for k, v in self.loc_files.items()})
        settings = self.config.get_settings()
        
        self.btn_run.config(state='disabled')
        self.prog.start(10)
        
        threading.Thread(
            target=run_logic_pandas, 
            args=(paths, settings, self.report_days.get(), self.log, self.done), 
            daemon=True
        ).start()

    def done(self, ok):
        self.prog.stop()
        self.btn_run.config(state='normal')
        if ok: 
            messagebox.showinfo("Success", "Analysis Complete. Report generated.")
        else: 
            messagebox.showerror("Failure", "Analysis failed. Please check the system log.")
