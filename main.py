import sys
from pathlib import Path
import tkinter as tk

from gui import ErasmusApp
from updater import check_for_updates

CURRENT_VERSION = 'v1.0.0'
GITHUB_USERNAME = 'l33ton'
GITHUB_REPO = 'Erasmus-Document-Generator'

if getattr(sys, 'frozen', False):
    base_dir = Path(sys.executable).parent
else:
    base_dir = Path(__file__).parent

if __name__ == "__main__":
    root = tk.Tk()
    app = ErasmusApp(root, base_dir)
    
    root.after(1000, lambda: check_for_updates(root, app.status_label, CURRENT_VERSION, GITHUB_USERNAME, GITHUB_REPO))
    
    root.mainloop()