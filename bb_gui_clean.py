"""
🔥 BLAZE BUDDY v3.0 - USER INTERFACE
Simple, Clean, Professional Interface
"""

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading
import queue
from datetime import datetime
from pathlib import Path

from bb_knowledge_base_clean import KnowledgeBase


class BlazeUI:
    """
    Clean, minimal interface for powerful knowledge base.
    One drop zone. Smart auto-detection. Beautiful feedback.
    """
    
    COLORS = {
        'primary': '#0066FF',
        'fire': '#FF6B35',
        'success': '#00D084',
        'dark_bg': '#0F1419',
        'card_bg': '#1A1F2E',
        'text_light': '#E8E8E8',
        'text_muted': '#888888',
        'border': '#2A3142',
    }
    
    def __init__(self, root):
        self.root = root
        self.root.title("🔥 BLAZE BUDDY v3.0")
        self.root.geometry("900x750")
        self.root.configure(bg=self.COLORS['dark_bg'])
        
        self.kb = KnowledgeBase()
        self.log_queue = queue.Queue()
        
        self.create_widgets()
        self._process_log_queue()
        
        # Auto-log startup
        self.log("System ready. Drop files to analyze.", 'info')
    
    def create_widgets(self):
        """Build UI"""
        
        # ===== HEADER =====
        header = tk.Frame(self.root, bg=self.COLORS['dark_bg'])
        header.pack(fill=tk.X, padx=20, pady=(20, 0))
        
        tk.Label(header, text="🔥 BLAZE BUDDY v3.0",
                font=('Segoe UI', 26, 'bold'),
                bg=self.COLORS['dark_bg'],
                fg=self.COLORS['primary']).pack(anchor='w')
        
        tk.Label(header, text="Cannabis Inventory Intelligence",
                font=('Segoe UI', 11),
                bg=self.COLORS['dark_bg'],
                fg=self.COLORS['fire']).pack(anchor='w')
        
        # ===== STATUS PANEL =====
        status = self.kb.get_status()
        
        status_frame = tk.Frame(self.root, bg=self.COLORS['card_bg'],
                               highlightbackground=self.COLORS['primary'],
                               highlightthickness=2)
        status_frame.pack(fill=tk.X, padx=20, pady=(15, 20))
        
        status_text = f"Knowledge Base Status:\n  • Products: {status['products']:,}\n  • Sales Records: {status['sales_records']:,}\n  • Confidence: {status['confidence']:.0%}"
        
        tk.Label(status_frame, text=status_text,
                font=('Consolas', 10),
                bg=self.COLORS['card_bg'],
                fg=self.COLORS['success'],
                justify=tk.LEFT).pack(anchor='w', padx=15, pady=15)
        
        # ===== DROP ZONE =====
        drop_container = tk.Frame(self.root, bg=self.COLORS['dark_bg'])
        drop_container.pack(fill=tk.BOTH, expand=True, padx=20, pady=(0, 15))
        
        drop_zone = tk.Frame(drop_container, bg=self.COLORS['card_bg'],
                            highlightbackground=self.COLORS['primary'],
                            highlightthickness=3)
        drop_zone.pack(fill=tk.BOTH, expand=True)
        
        inner = tk.Frame(drop_zone, bg=self.COLORS['card_bg'])
        inner.pack(fill=tk.BOTH, expand=True, padx=40, pady=40)
        
        tk.Label(inner, text="📁",
                font=('Arial', 60),
                bg=self.COLORS['card_bg']).pack(pady=(20, 0))
        
        tk.Label(inner, text="Drop Files Here",
                font=('Segoe UI', 18, 'bold'),
                bg=self.COLORS['card_bg'],
                fg=self.COLORS['text_light']).pack(pady=10)
        
        tk.Label(inner, text="Government Catalog, Inventory, Sales, or POs",
                font=('Segoe UI', 11),
                bg=self.COLORS['card_bg'],
                fg=self.COLORS['text_muted']).pack(pady=5)
        
        tk.Label(inner, text="Auto-detects type and learns",
                font=('Segoe UI', 9, 'italic'),
                bg=self.COLORS['card_bg'],
                fg=self.COLORS['text_muted']).pack(pady=(20, 0))
        
        ttk.Button(inner, text="📂 Browse Files",
                  command=self.browse_files).pack(pady=15)
        
        # ===== LOG AREA =====
        log_label = tk.Label(self.root, text="System Log",
                            font=('Segoe UI', 10, 'bold'),
                            bg=self.COLORS['dark_bg'],
                            fg=self.COLORS['primary'])
        log_label.pack(anchor='w', padx=20, pady=(10, 5))
        
        self.log_text = tk.Text(self.root, height=8,
                               font=('Consolas', 9),
                               bg=self.COLORS['dark_bg'],
                               fg=self.COLORS['text_light'],
                               insertbackground=self.COLORS['primary'],
                               highlightthickness=0,
                               border=0)
        self.log_text.pack(fill=tk.BOTH, expand=True, padx=20, pady=(0, 20))
        self.log_text.config(state='disabled')
        
        # Configure tags
        self.log_text.tag_config('success', foreground=self.COLORS['success'])
        self.log_text.tag_config('error', foreground='#FF3B30')
        self.log_text.tag_config('info', foreground=self.COLORS['primary'])
    
    def browse_files(self):
        """Browse and import files"""
        files = filedialog.askopenfilenames(
            filetypes=[
                ("All Supported", "*.csv *.xlsx *.xlsm *.xls"),
                ("CSV", "*.csv"),
                ("Excel", "*.xlsx *.xlsm *.xls"),
                ("All", "*.*")
            ]
        )
        
        for filepath in files:
            self.ingest_file(filepath)
    
    def ingest_file(self, filepath: str):
        """Ingest file in background"""
        threading.Thread(
            target=self._ingest_thread,
            args=(filepath,),
            daemon=True
        ).start()
    
    def _ingest_thread(self, filepath: str):
        """Background ingestion"""
        try:
            self.log(f"Analyzing: {Path(filepath).name}...", 'info')
            result = self.kb.ingest_file(filepath)
            
            if result['success']:
                self.log(result['message'], 'success')
                # Update status
                status = self.kb.get_status()
                self.log(f"Knowledge Base: {status['products']:,} products, {status['confidence']:.0%} confidence", 'info')
            else:
                self.log(result['message'], 'error')
        
        except Exception as e:
            self.log(f"Error: {str(e)}", 'error')
    
    def log(self, msg: str, tag: str = 'info'):
        """Queue log message"""
        self.log_queue.put((msg, tag))
    
    def _process_log_queue(self):
        """Process log messages"""
        try:
            while True:
                msg, tag = self.log_queue.get_nowait()
                self.log_text.config(state='normal')
                ts = datetime.now().strftime('%H:%M:%S')
                self.log_text.insert(tk.END, f"[{ts}] {msg}\n", tag)
                self.log_text.see(tk.END)
                self.log_text.config(state='disabled')
        except queue.Empty:
            pass
        
        self.root.after(100, self._process_log_queue)


if __name__ == "__main__":
    root = tk.Tk()
    app = BlazeUI(root)
    root.mainloop()
