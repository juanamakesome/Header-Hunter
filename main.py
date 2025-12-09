"""
Header Hunter v5.0 - Main Entry Point
Modular Inventory Management System
"""
import tkinter as tk
from hh_gui import GreenlineApp

if __name__ == "__main__":
    """
    Application entry point. Creates Tkinter root window and initializes the GUI.
    """
    root = tk.Tk()
    app = GreenlineApp(root)
    root.mainloop()