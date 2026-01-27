"""
Header Hunter v8.0 - Drag & Drop Interface
Only requires Inventory + Sales (AGLC is optional reference)
"""
import customtkinter as ctk
from tkinter import filedialog, messagebox
import threading
import queue
import os
from datetime import datetime

# IMPORTS
try:
    from hh_utils import APP_TITLE, load_config, save_config
    from hh_logic import run_logic_pandas
except ImportError as e:
    # Logging not available yet - use print for critical startup errors
    print(f"CRITICAL: Missing dependencies: {e}")
    print("pip install customtkinter pandas xlsxwriter openpyxl")
    exit(1)

# Try to import tkinterdnd2 for drag-and-drop
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    DND_AVAILABLE = True
except ImportError:
    DND_AVAILABLE = False
    # Logging not available yet - use print for optional feature warnings
    print("Note: tkinterdnd2 not installed. Drag-and-drop disabled.")
    print("Install with: pip install tkinterdnd2")

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("dark-blue")

# Create class with conditional drag-and-drop support
if DND_AVAILABLE:
    class HeaderHunterCockpit(ctk.CTk, TkinterDnD.DnDWrapper):
        """Cockpit with drag-and-drop support."""
        
        def __init__(self):
            ctk.CTk.__init__(self)
            self.TkdndVersion = TkinterDnD._require(self)
            self._init_common()
else:
    class HeaderHunterCockpit(ctk.CTk):
        """Cockpit without drag-and-drop."""
        
        def __init__(self):
            super().__init__()
            self._init_common()

def _init_common(self):
    """Common initialization."""
    self.title("🎯 Header Hunter | Cockpit")
    self.geometry("800x700")
    
    # Set window icon
    try:
        import sys
        if getattr(sys, 'frozen', False):
            # Running as compiled executable
            base_path = sys._MEIPASS
        else:
            # Running as script
            base_path = os.path.dirname(os.path.abspath(__file__))
        icon_path = os.path.join(base_path, 'icon.ico')
        if os.path.exists(icon_path):
            self.iconbitmap(icon_path)
    except Exception:
        pass  # Icon is optional
    
    self.files = {
        k: None for k in ['inventory', 'sales', 'aglc', 'po', 'hill', 'valley', 'jasper']
    }
    
    self.config_data = load_config()
    
    # Initialize PO Destination state
    saved_dest = self.config_data.get('settings', {}).get('po_destination', 'J')
    dest_map = {'H': 'Hill', 'V': 'Valley', 'J': 'Jasper'}
    self.po_dest_var = ctk.StringVar(value=dest_map.get(saved_dest, 'Jasper'))
    
    self.log_queue = queue.Queue()
    
    self.create_ui()
    self._process_log_queue()
    
    # Enable drag-and-drop on entire window if available
    if DND_AVAILABLE:
        try:
            self.drop_target_register(DND_FILES)
            self.dnd_bind('<<Drop>>', self.on_file_drop)
            self.dnd_bind('<<DragEnter>>', self.on_drag_enter)
            self.dnd_bind('<<DragLeave>>', self.on_drag_leave)
            self.log("✅ Drag-and-drop enabled")
        except Exception as e:
            self.log(f"⚠️ Drag-and-drop setup failed: {e}")

def create_ui(self):
    """Create the UI."""
    # HEADER
    header = ctk.CTkFrame(self, fg_color="#1a1a2e", corner_radius=0)
    header.pack(fill="x", pady=0)
    
    ctk.CTkLabel(
        header, 
        text="🎯 HEADER HUNTER", 
        font=ctk.CTkFont(size=28, weight="bold"),
        text_color="#00ff88"
    ).pack(pady=20)
    
    # MAIN CONTAINER
    main = ctk.CTkFrame(self, fg_color="transparent")
    main.pack(fill="both", expand=True, padx=30, pady=20)
    
    # DROP ZONE
    self.drop_zone = ctk.CTkFrame(
        main, 
        fg_color="#0f3460",
        border_width=3,
        border_color="#00ff88",
        corner_radius=20,
        height=300
    )
    self.drop_zone.pack(fill="both", expand=True, pady=(0, 15))
    self.drop_zone.pack_propagate(False)
    
    # Drop content
    drop_content = ctk.CTkFrame(self.drop_zone, fg_color="transparent")
    drop_content.place(relx=0.5, rely=0.3, anchor="center")
    
    ctk.CTkLabel(
        drop_content,
        text="📂",
        font=ctk.CTkFont(size=64)
    ).pack()
    
    ctk.CTkLabel(
        drop_content,
        text="DRAG FILES HERE\nor click to browse",
        font=ctk.CTkFont(size=18, weight="bold"),
        text_color="#00ff88"
    ).pack(pady=15)
    
    # Make clickable
    self.drop_zone.bind("<Button-1>", lambda e: self.browse_folder())
    drop_content.bind("<Button-1>", lambda e: self.browse_folder())
    
    # FILE STATUS
    status_frame = ctk.CTkFrame(self.drop_zone, fg_color="#16213e", corner_radius=10)
    status_frame.place(relx=0.5, rely=0.75, anchor="center", relwidth=0.85)
    
    self.file_labels = {}
    file_list = [
        ('inventory', 'Inventory', True),  # Required
        ('sales', 'Sales', True),  # Required
        ('aglc', 'AGLC Manual (Optional)', False)  # Optional
    ]
    
    for key, label, required in file_list:
        row = ctk.CTkFrame(status_frame, fg_color="transparent")
        row.pack(fill="x", padx=15, pady=4)
        
        icon_lbl = ctk.CTkLabel(row, text="⚪", font=ctk.CTkFont(size=14), width=25)
        icon_lbl.pack(side="left", padx=5)
        
        label_text = f"{'*' if required else ''} {label}"
        text_lbl = ctk.CTkLabel(
            row, 
            text=label_text,
            font=ctk.CTkFont(size=12),
            anchor="w"
        )
        text_lbl.pack(side="left", fill="x", expand=True)
        
        self.file_labels[key] = (icon_lbl, text_lbl)
    
    # PO DESTINATION SELECTOR
    dest_frame = ctk.CTkFrame(main, fg_color="transparent")
    dest_frame.pack(fill="x", pady=(0, 15))
    
    ctk.CTkLabel(
        dest_frame, 
        text="📦 PO Destination:", 
        font=ctk.CTkFont(size=14, weight="bold")
    ).pack(side="left", padx=(0, 10))
    
    self.dest_selector = ctk.CTkSegmentedButton(
        dest_frame,
        values=["Hill", "Valley", "Jasper"],
        variable=self.po_dest_var,
        command=lambda v: self.save_paths(),
        height=35,
        font=ctk.CTkFont(size=12)
    )
    self.dest_selector.pack(side="left", fill="x", expand=True)
    
    # GENERATE BUTTON
    self.gen_btn = ctk.CTkButton(
        main,
        text="🚀 GENERATE REPORT",
        command=self.generate_report,
        font=ctk.CTkFont(size=18, weight="bold"),
        height=50,
        fg_color="#00cc66",
        hover_color="#00ff88",
        text_color="#000000",
        corner_radius=10
    )
    self.gen_btn.pack(fill="x", pady=(0, 10))
    
    # PROGRESS BAR
    self.progress = ctk.CTkProgressBar(main, height=6, corner_radius=3)
    self.progress.pack(fill="x")
    self.progress.set(0)
    
    # LOG (initially hidden)
    self.log_visible = False
    self.log_text = ctk.CTkTextbox(
        main, 
        height=0,
        font=("Consolas", 10),
        fg_color="#0f0f23",
        text_color="#00ff88"
    )

def on_file_drop(self, event):
    """Handle file drop event."""
    if not DND_AVAILABLE:
        return
    
    try:
        files = self.tk.splitlist(event.data)
        self.log(f"📥 Received {len(files)} file(s)")
        
        # Log all files for debugging
        for f in files:
            self.log(f"  📄 {os.path.basename(f)}")
        
        assigned_files = []
        for filepath in files:
            # Normalize path
            filepath = filepath.replace('{', '').replace('}', '')
            if not os.path.exists(filepath):
                continue
                
            filename = os.path.basename(filepath).lower()
            assigned = False
            
            # Auto-assign based on filename
            if 'inventory' in filename:
                self.assign_file('inventory', filepath)
                assigned = True
            elif 'sales' in filename or 'product-sales' in filename:
                self.assign_file('sales', filepath)
                assigned = True
            elif ('manual' in filename or 'aglc' in filename or 'cannabisretailers' in filename) and not self.files['aglc']:
                self.assign_file('aglc', filepath)
                assigned = True
            elif ('purchase' in filename or 'po' in filename or filename.startswith('p0')) and not self.files['po']:
                self.assign_file('po', filepath)
                assigned = True
            elif 'transfer' in filename or 'trans' in filename:
                # Transfer files with "transfer" or "trans" in name
                # Auto-assign to location if specified in filename, otherwise assign to first available slot
                # The logic will parse Source/Dest columns from the file itself
                if 'hill' in filename and not self.files['hill']:
                    self.assign_file('hill', filepath)
                    assigned = True
                elif 'valley' in filename and not self.files['valley']:
                    self.assign_file('valley', filepath)
                    assigned = True
                elif 'jasper' in filename and not self.files['jasper']:
                    self.assign_file('jasper', filepath)
                    assigned = True
                else:
                    # Transfer file without location name - assign to first available slot
                    # The logic will parse Source/Dest columns from the file
                    if not self.files['hill']:
                        self.assign_file('hill', filepath)
                        assigned = True
                    elif not self.files['valley']:
                        self.assign_file('valley', filepath)
                        assigned = True
                    elif not self.files['jasper']:
                        self.assign_file('jasper', filepath)
                        assigned = True
                    else:
                        self.log(f"  ⚠️ Transfer file skipped (all transfer slots full): {os.path.basename(filepath)}")
            elif not assigned:
                # Check for location names without "transfer" keyword
                # This handles files like "hill.csv", "valley-inventory.csv", etc.
                if 'hill' in filename and not self.files['hill'] and filename.endswith('.csv'):
                    self.assign_file('hill', filepath)
                    assigned = True
                elif 'valley' in filename and not self.files['valley'] and filename.endswith('.csv'):
                    self.assign_file('valley', filepath)
                    assigned = True
                elif 'jasper' in filename and not self.files['jasper'] and filename.endswith('.csv'):
                    self.assign_file('jasper', filepath)
                    assigned = True
            
            if not assigned:
                # Log unrecognized files for debugging
                self.log(f"  ⚠️ Unrecognized file: {os.path.basename(filepath)}")
        
        self.save_paths()
        self.drop_zone.configure(border_color="#00ff88", fg_color="#0f3460")
        self.after(300, lambda: self.drop_zone.configure(border_color="#00ff88"))
    except Exception as e:
        self.log(f"⚠️ Error processing dropped files: {e}")

def on_drag_enter(self, event):
    """Visual feedback when dragging."""
    if DND_AVAILABLE:
        self.drop_zone.configure(border_color="#ffcc00", fg_color="#1a4d2e")

def on_drag_leave(self, event):
    """Reset visual feedback."""
    if DND_AVAILABLE:
        self.drop_zone.configure(border_color="#00ff88", fg_color="#0f3460")

def browse_folder(self):
    """Browse for folder."""
    folder = filedialog.askdirectory(title="Select folder with data files")
    if not folder:
        return
    
    self.show_log()
    self.log(f"📁 Scanning {os.path.basename(folder)}...")
    
    files = [
        os.path.join(folder, f) 
        for f in os.listdir(folder) 
        if f.lower().endswith(('.csv', '.xlsx', '.xlsm')) and not f.startswith('~$')
    ]
    
    # Auto-detect and assign
    for f in files:
        name = os.path.basename(f).lower()
        assigned = False
        
        if 'inventory' in name:
            self.assign_file('inventory', f)
            assigned = True
        elif 'sales' in name or 'product-sales' in name:
            self.assign_file('sales', f)
            assigned = True
        elif ('manual' in name or 'aglc' in name or 'cannabisretailers' in name) and not self.files['aglc']:
            self.assign_file('aglc', f)
            assigned = True
        elif ('purchase' in name or 'po' in name or name.startswith('p0')) and not self.files['po']:
            self.assign_file('po', f)
            assigned = True
        elif 'transfer' in name or 'trans' in name:
            # Transfer files with "transfer" or "trans" in name
            # Auto-assign to location if specified in filename, otherwise assign to first available slot
            # The logic will parse Source/Dest columns from the file itself
            if 'hill' in name and not self.files['hill']:
                self.assign_file('hill', f)
                assigned = True
            elif 'valley' in name and not self.files['valley']:
                self.assign_file('valley', f)
                assigned = True
            elif 'jasper' in name and not self.files['jasper']:
                self.assign_file('jasper', f)
                assigned = True
            else:
                # Transfer file without location name - assign to first available slot
                # The logic will parse Source/Dest columns from the file
                if not self.files['hill']:
                    self.assign_file('hill', f)
                    assigned = True
                elif not self.files['valley']:
                    self.assign_file('valley', f)
                    assigned = True
                elif not self.files['jasper']:
                    self.assign_file('jasper', f)
                    assigned = True
                else:
                    self.log(f"  ⚠️ Transfer file skipped (all transfer slots full): {os.path.basename(f)}")
        elif not assigned:
            # Check for location names without "transfer" keyword
            if 'hill' in name and not self.files['hill'] and name.endswith('.csv'):
                self.assign_file('hill', f)
                assigned = True
            elif 'valley' in name and not self.files['valley'] and name.endswith('.csv'):
                self.assign_file('valley', f)
                assigned = True
            elif 'jasper' in name and not self.files['jasper'] and name.endswith('.csv'):
                self.assign_file('jasper', f)
                assigned = True
        
        if not assigned:
            self.log(f"  ⚠️ Unrecognized file: {os.path.basename(f)}")

def assign_file(self, key, filepath):
    """Assign file and update UI."""
    self.files[key] = filepath
    
    if key in self.file_labels:
        icon_lbl, text_lbl = self.file_labels[key]
        icon_lbl.configure(text="✅", text_color="#00ff88")
        filename = os.path.basename(filepath)
        # Update label to show filename
        base_label = text_lbl.cget('text').split(' • ')[0]
        text_lbl.configure(text=f"{base_label} • {filename[:30]}")
    
    self.log(f"  ✓ {key.title()}: {os.path.basename(filepath)}")

def show_log(self):
    """Show log section."""
    if not self.log_visible:
        self.log_text.pack(fill="both", expand=True, pady=(10, 0))
        self.log_text.configure(height=120)
        self.log_visible = True

def log(self, msg):
    """Add log message."""
    self.log_queue.put(msg)

def _process_log_queue(self):
    """Process log messages."""
    try:
        while True:
            msg = self.log_queue.get_nowait()
            ts = datetime.now().strftime('%H:%M:%S')
            self.log_text.insert("end", f"[{ts}] {msg}\n")
            self.log_text.see("end")
    except queue.Empty:
        pass
    self.after(100, self._process_log_queue)

def generate_report(self):
    """Generate Excel report."""
    # Only require Inventory and Sales
    required = ['inventory', 'sales']
    missing = [k for k in required if not self.files[k]]
    
    if missing:
        messagebox.showerror(
            "Missing Required Files",
            f"Required files:\n• {', '.join(m.title() for m in missing)}\n\nAGLC Manual is optional (reference only)."
        )
        return
    
    # Get PO destination from UI
    dest_display = self.po_dest_var.get()
    dest_map = {'Hill': 'H', 'Valley': 'V', 'Jasper': 'J'}
    po_destination = dest_map.get(dest_display, 'J')
    
    self.log(f"📦 PO Destination: {po_destination} ({dest_display})")
    
    self.show_log()
    self.gen_btn.configure(state="disabled", text="⏳ Generating...")
    self.progress.set(0)
    self.progress.start()
    self.log("🚀 Starting analysis...")
    
    # Run in thread
    paths = {k: v for k, v in self.files.items() if v}  # Only include files that exist
    settings = self.config_data.get('settings', {})
    
    # Add PO destination to settings
    if po_destination:
        settings['po_destination'] = po_destination
    
    t = threading.Thread(
        target=run_logic_pandas,
        args=(paths, settings, "30", self.log, self.on_complete),
        daemon=True
    )
    t.start()

def on_complete(self, success):
    """Handle completion."""
    self.progress.stop()
    self.progress.set(1 if success else 0)
    self.gen_btn.configure(state="normal", text="🚀 GENERATE REPORT")
    
    if success:
        self.log("✅ Report generated and opened!")
        messagebox.showinfo("✅ Complete", "Report generated and opened!")
    else:
        self.log("❌ Generation failed - check log above")
        messagebox.showerror("❌ Error", "Generation failed. Check log for details.")

def save_paths(self):
    """Save file paths to config."""
    if 'paths' not in self.config_data:
        self.config_data['paths'] = {}
    # Convert None to empty string for JSON
    self.config_data['paths'].update({k: (v or '') for k, v in self.files.items()})
    
    # Save PO Destination preference
    if 'settings' not in self.config_data:
        self.config_data['settings'] = {}
    
    dest_display = self.po_dest_var.get()
    dest_map = {'Hill': 'H', 'Valley': 'V', 'Jasper': 'J'}
    self.config_data['settings']['po_destination'] = dest_map.get(dest_display, 'J')
    
    save_config(self.config_data)

# Attach all methods to both class versions
HeaderHunterCockpit._init_common = _init_common
HeaderHunterCockpit.create_ui = create_ui
HeaderHunterCockpit.on_file_drop = on_file_drop
HeaderHunterCockpit.on_drag_enter = on_drag_enter
HeaderHunterCockpit.on_drag_leave = on_drag_leave
HeaderHunterCockpit.browse_folder = browse_folder
HeaderHunterCockpit.assign_file = assign_file
HeaderHunterCockpit.show_log = show_log
HeaderHunterCockpit.log = log
HeaderHunterCockpit._process_log_queue = _process_log_queue
HeaderHunterCockpit.generate_report = generate_report
HeaderHunterCockpit.on_complete = on_complete
HeaderHunterCockpit.save_paths = save_paths
