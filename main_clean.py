"""
🔥 BLAZE BUDDY v3.0
Cannabis Inventory Intelligence System
Main Entry Point
"""

import tkinter as tk
from bb_gui_clean import BlazeUI


def main():
    """Launch the application"""
    root = tk.Tk()
    app = BlazeUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
