"""
Header Hunter v5.4 - GUI Module
Added: Configuration for Downloads/Ingest Folder
"""
import tkinter as tk
from tkinter import ttk, filedialog, scrolledtext, messagebox
import threading
import queue
import os
from datetime import datetime
from hh_utils import APP_TITLE, DEFAULT_SETTINGS, resource_path, load_config, save_config
from hh_logic import run_logic_pandas
import hh_ingest


class GreenlineApp:
    
    def __init__(self, root):
        self.root = root
        self.root.title(APP_TITLE)
        self.root.geometry("850x900")
        
        try:
            icon_path = resource_path("icon.ico")
            if os.path.exists(icon_path):
                self.root.iconbitmap(icon_path)
        except: pass
        
        style = ttk.Style()
        style.theme_use('clam')
        style.configure("TButton", padding=6, font=('Segoe UI', 10))
        style.configure("Header.TLabel", font=('Segoe UI', 18, 'bold'), foreground="#2E8B57")
        style.configure("Sub.TLabel", font=('Segoe UI', 9), foreground="#666666")
        
        self.files = {
            'inventory': tk.StringVar(), 'sales': tk.StringVar(), 'po': tk.StringVar(),
            'aglc': tk.StringVar(), 'hill': tk.StringVar(), 'valley': tk.StringVar(), 'jasper': tk.StringVar()
        }
        self.report_days = tk.StringVar(value="30")
        self.config_data = load_config()
        
        paths = self.config_data.get('paths', {})
        for k, v in paths.items():
            if k in self.files: self.files[k].set(v)
        
        self.log_queue = queue.Queue()
        self.create_widgets()
        self._process_log_queue()
    
    def create_widgets(self):
        header = ttk.Frame(self.root, padding="20")
        header.pack(fill=tk.X)
        
        try:
            logo_path = resource_path("logo.png")
            if os.path.exists(logo_path):
                img = tk.PhotoImage(file=logo_path)
                if img.width() > 100: img = img.subsample(int(img.width() / 100))
                self.logo_img = img
                ttk.Label(header, image=self.logo_img).pack(side=tk.LEFT, padx=10)
        except: pass
        
        title_box = ttk.Frame(header)
        title_box.pack(side=tk.LEFT, padx=10)
        ttk.Label(title_box, text=APP_TITLE, style="Header.TLabel").pack(anchor='w')
        ttk.Label(title_box, text="Modular Intelligence System", style="Sub.TLabel").pack(anchor='w')
        
        btn_box = ttk.Frame(header)
        btn_box.pack(side=tk.RIGHT)
        
        # Buttons
        ttk.Button(btn_box, text="📥 Ingest Data", command=self.run_ingest_thread).pack(fill=tk.X, pady=2)
        ttk.Button(btn_box, text="📂 Auto-Load", command=self.auto_load_folder).pack(fill=tk.X, pady=2)
        ttk.Button(btn_box, text="⚙️ Settings", command=self.open_settings).pack(fill=tk.X, pady=2)
        
        main_frame = ttk.Frame(self.root, padding=20)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        f_params = ttk.LabelFrame(main_frame, text="Run Parameters", padding=10)
        f_params.pack(fill=tk.X, pady=5)
        ttk.Label(f_params, text="Days of Sales Data:").pack(side=tk.LEFT)
        ttk.Entry(f_params, textvariable=self.report_days, width=5).pack(side=tk.LEFT, padx=5)
        ttk.Label(f_params, text="(Used for velocity calc)").pack(side=tk.LEFT)
        
        lf_req = ttk.LabelFrame(main_frame, text="Required Files", padding=10)
        lf_req.pack(fill=tk.X, pady=5)
        self.make_file_row(lf_req, "Inventory Export:", 'inventory')
        self.make_file_row(lf_req, "Sales Data:", 'sales')
        self.make_file_row(lf_req, "Purchase Order:", 'po')
        self.make_file_row(lf_req, "AGLC Manual Form:", 'aglc')
        
        lf_trans = ttk.LabelFrame(main_frame, text="Pending Transfers", padding=10)
        lf_trans.pack(fill=tk.X, pady=5)
        self.make_file_row(lf_trans, "To Hill:", 'hill')
        self.make_file_row(lf_trans, "To Valley:", 'valley')
        self.make_file_row(lf_trans, "To Jasper:", 'jasper')
        
        self.progress = ttk.Progressbar(main_frame, orient="horizontal", mode="indeterminate")
        self.progress.pack(fill=tk.X, pady=(20, 5))
        
        self.run_btn = ttk.Button(main_frame, text="Kolwalski, Analysis! 🐧", command=self.start_processing)
        self.run_btn.pack(fill=tk.X, ipady=10)
        
        log_lf = ttk.LabelFrame(main_frame, text="System Log", padding=5)
        log_lf.pack(fill=tk.BOTH, expand=True, pady=10)
        self.log_area = scrolledtext.ScrolledText(log_lf, height=8, font=('Consolas', 9), state='disabled')
        self.log_area.pack(fill=tk.BOTH, expand=True)
    
    def make_file_row(self, parent, label, key):
        f = ttk.Frame(parent)
        f.pack(fill=tk.X, pady=2)
        ttk.Label(f, text=label, width=18).pack(side=tk.LEFT)
        ttk.Entry(f, textvariable=self.files[key]).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        ttk.Button(f, text="...", width=4, command=lambda: self.browse(key)).pack(side=tk.RIGHT)
    
    def browse(self, key):
        path = filedialog.askopenfilename(filetypes=[("Data Files", "*.csv *.xlsx *.xlsm"), ("All", "*.*")])
        if path:
            self.files[key].set(path)
            self.save_paths()
    
    def log(self, msg):
        self.log_queue.put(msg)
    
    def _process_log_queue(self):
        try:
            while True:
                msg = self.log_queue.get_nowait()
                self.log_area.config(state='normal')
                ts = datetime.now().strftime('%H:%M:%S')
                self.log_area.insert(tk.END, f"[{ts}] {msg}\n")
                self.log_area.see(tk.END)
                self.log_area.config(state='disabled')
        except queue.Empty: pass
        self.root.after(100, self._process_log_queue)
    
    def save_paths(self):
        self.config_data['paths'] = {k: v.get() for k, v in self.files.items()}
        save_config(self.config_data)
    
    def peek_transfer_location(self, filepath):
        try:
            with open(filepath, 'r', encoding='utf-8', errors='ignore') as f:
                content = f.read(4096).lower()
                scores = {'hill': content.count('hill'), 'valley': content.count('valley'), 'jasper': content.count('jasper')}
                best_loc = max(scores, key=scores.get)
                if scores[best_loc] > 0: return best_loc
                return None
        except: return None
    
    def auto_load_folder(self):
        folder_path = filedialog.askdirectory(title="Select Download Folder")
        if not folder_path: return
        
        try:
            all_files = [os.path.join(folder_path, f) for f in os.listdir(folder_path) if f.lower().endswith(('.csv', '.xlsx', '.xlsm')) and not f.startswith('~$')]
            all_files.sort(key=os.path.getmtime, reverse=True)
        except OSError as e:
            messagebox.showerror("Error", f"Could not read folder: {e}")
            return
        
        self.log(f"🕵️ Scanning {len(all_files)} files...")
        assigned = {'inventory': 0, 'sales': 0, 'po': 0, 'aglc': 0, 'hill': 0, 'valley': 0, 'jasper': 0}
        found_types = []
        
        for full_path in all_files:
            filename = os.path.basename(full_path).lower()
            if "inventory-export" in filename and 'inventory' not in found_types:
                self.files['inventory'].set(full_path); assigned['inventory'] += 1; found_types.append('inventory'); self.log(f" -> Found Inventory: {filename}")
            elif "product-sales" in filename and 'sales' not in found_types:
                self.files['sales'].set(full_path); assigned['sales'] += 1; found_types.append('sales'); self.log(f" -> Found Sales: {filename}")
            elif "purchase-order" in filename and 'po' not in found_types:
                self.files['po'].set(full_path); assigned['po'] += 1; found_types.append('po'); self.log(f" -> Found PO: {filename}")
            elif ("manual" in filename or "retailers" in filename) and "order" in filename and 'aglc' not in found_types:
                self.files['aglc'].set(full_path); assigned['aglc'] += 1; found_types.append('aglc'); self.log(f" -> Found AGLC: {filename}")
            elif "transfer" in filename:
                loc = self.peek_transfer_location(full_path)
                if loc and loc not in found_types:
                    self.files[loc].set(full_path); assigned[loc] += 1; found_types.append(loc); self.log(f" -> Found {loc.title()} Transfer: {filename}")
        
        self.save_paths()
        messagebox.showinfo("Scan Report", f"Found {sum(assigned.values())} matching files.")
    
    def run_ingest_thread(self):
        """Runs the Ingest process in a separate thread."""
        self.log("--- 🧠 Starting Memory Bank Ingest ---")
        t = threading.Thread(target=self._ingest_wrapper, daemon=True)
        t.start()

    def _ingest_wrapper(self):
        try:
            config = load_config()
            hist_path = config.get('settings', {}).get('history_folder', '')
            if not hist_path:
                self.log("❌ ERROR: Memory Bank folder not configured! Go to Settings.")
                return
            hh_ingest.update_memory_bank(self.log)
            self.log("--- Ingest Complete ---")
        except Exception as e:
            self.log(f"❌ Ingest Error: {e}")

    def open_settings(self):
        top = tk.Toplevel(self.root)
        top.title("Settings")
        top.geometry("600x750")
        self.entry_widgets = {}
        
        def make_inp(parent, label, key, val, row):
            ttk.Label(parent, text=label).grid(row=row, column=0, sticky='w', pady=2)
            e = ttk.Entry(parent)
            e.insert(0, str(val))
            e.grid(row=row, column=1, sticky='ew', padx=5, pady=2)
            self.entry_widgets[key] = e
        
        curr_sets = self.config_data.get('settings', DEFAULT_SETTINGS)
        
        # Cannabis Rules
        lf_can = ttk.LabelFrame(top, text="Cannabis Logic", padding=10)
        lf_can.pack(fill=tk.X, padx=10, pady=5)
        c_rules = curr_sets.get('cannabis_logic', DEFAULT_SETTINGS['cannabis_logic'])
        make_inp(lf_can, "Hot Velocity:", 'c_hot', c_rules['hot_velocity'], 0)
        make_inp(lf_can, "Reorder Point:", 'c_reorder', c_rules['reorder_point'], 1)
        make_inp(lf_can, "Target WOS:", 'c_target', c_rules.get('target_wos', 4.0), 2)
        make_inp(lf_can, "Dead WOS:", 'c_dead_wos', c_rules.get('dead_wos', 26), 3)
        make_inp(lf_can, "Dead Stock:", 'c_dead_oh', c_rules.get('dead_on_hand', 5), 4)
        
        # Accessory Rules
        lf_acc = ttk.LabelFrame(top, text="Accessory Logic", padding=10)
        lf_acc.pack(fill=tk.X, padx=10, pady=5)
        a_rules = curr_sets.get('accessory_logic', DEFAULT_SETTINGS['accessory_logic'])
        make_inp(lf_acc, "Hot Velocity:", 'a_hot', a_rules['hot_velocity'], 0)
        make_inp(lf_acc, "Reorder Point:", 'a_reorder', a_rules['reorder_point'], 1)
        make_inp(lf_acc, "Target WOS:", 'a_target', a_rules.get('target_wos', 8.0), 2)
        make_inp(lf_acc, "Dead WOS:", 'a_dead_wos', a_rules.get('dead_wos', 52), 3)
        make_inp(lf_acc, "Dead Stock:", 'a_dead_oh', a_rules.get('dead_on_hand', 3), 4)

        # History Settings (Updated)
        lf_hist = ttk.LabelFrame(top, text="Workflow Settings", padding=10)
        lf_hist.pack(fill=tk.X, padx=10, pady=5)
        lf_hist.columnconfigure(1, weight=1)
        
        # 1. Memory Bank
        curr_hist_folder = curr_sets.get('history_folder', '')
        ttk.Label(lf_hist, text="Memory Bank Folder:").grid(row=0, column=0, sticky='w', pady=2)
        hist_entry = ttk.Entry(lf_hist)
        hist_entry.insert(0, curr_hist_folder)
        hist_entry.grid(row=0, column=1, sticky='ew', padx=5, pady=2)
        
        def browse_history():
            folder = filedialog.askdirectory(title="Select MEMORYBANK Folder")
            if folder:
                hist_entry.delete(0, tk.END)
                hist_entry.insert(0, folder)
        ttk.Button(lf_hist, text="Browse", command=browse_history).grid(row=0, column=2, padx=5)
        ttk.Label(lf_hist, text="(Storage for Master DB)", style="Sub.TLabel").grid(row=1, column=0, columnspan=3, sticky='w', pady=(0, 10))

        # 2. Ingest Source
        curr_ingest_folder = curr_sets.get('ingest_folder', '')
        ttk.Label(lf_hist, text="Downloads Folder:").grid(row=2, column=0, sticky='w', pady=2)
        ingest_entry = ttk.Entry(lf_hist)
        ingest_entry.insert(0, curr_ingest_folder)
        ingest_entry.grid(row=2, column=1, sticky='ew', padx=5, pady=2)
        
        def browse_ingest():
            folder = filedialog.askdirectory(title="Select Downloads Folder")
            if folder:
                ingest_entry.delete(0, tk.END)
                ingest_entry.insert(0, folder)
        ttk.Button(lf_hist, text="Browse", command=browse_ingest).grid(row=2, column=2, padx=5)
        ttk.Label(lf_hist, text="(Leave empty to auto-detect Downloads)", style="Sub.TLabel").grid(row=3, column=0, columnspan=3, sticky='w', pady=2)

        def save_close():
            try:
                new_settings = self.config_data.get('settings', DEFAULT_SETTINGS)
                # Save Logic
                new_settings['cannabis_logic'].update({
                    'hot_velocity': float(self.entry_widgets['c_hot'].get()),
                    'reorder_point': float(self.entry_widgets['c_reorder'].get()),
                    'target_wos': float(self.entry_widgets['c_target'].get()),
                    'dead_wos': float(self.entry_widgets['c_dead_wos'].get()),
                    'dead_on_hand': float(self.entry_widgets['c_dead_oh'].get())
                })
                new_settings['accessory_logic'].update({
                    'hot_velocity': float(self.entry_widgets['a_hot'].get()),
                    'reorder_point': float(self.entry_widgets['a_reorder'].get()),
                    'target_wos': float(self.entry_widgets['a_target'].get()),
                    'dead_wos': float(self.entry_widgets['a_dead_wos'].get()),
                    'dead_on_hand': float(self.entry_widgets['a_dead_oh'].get())
                })
                # Save Paths
                new_settings['history_folder'] = hist_entry.get()
                new_settings['ingest_folder'] = ingest_entry.get()
                
                self.config_data['settings'] = new_settings
                save_config(self.config_data)
                top.destroy()
            except ValueError: messagebox.showerror("Error", "Values must be numbers.")
        
        ttk.Button(top, text="Save Configuration", command=save_close).pack(pady=10, fill=tk.X, padx=20)
    
    def start_processing(self):
        if not self.files['inventory'].get() or not self.files['sales'].get():
            messagebox.showerror("Missing Files", "Inventory and Sales files are mandatory.")
            return
        
        self.run_btn.config(state='disabled')
        self.progress.start(10)
        self.log_area.config(state='normal')
        self.log_area.delete(1.0, tk.END)
        self.log_area.config(state='disabled')
        
        paths = {k: v.get() for k, v in self.files.items()}
        settings = self.config_data.get('settings', DEFAULT_SETTINGS)
        report_days = self.report_days.get()
        
        t = threading.Thread(
            target=run_logic_pandas,
            args=(paths, settings, report_days, self.log, self.on_finish),
            daemon=True
        )
        t.start()
    
    def on_finish(self, success):
        self.progress.stop()
        self.run_btn.config(state='normal')
        if success: messagebox.showinfo("Success", "Analysis Complete!")
        else: messagebox.showerror("Error", "Analysis failed. Check log.")
