import sys
from pathlib import Path
import tkinter as tk

from gui import ErasmusApp
from updater import check_for_updates

CURRENT_VERSION = 'v1.0.1'
GITHUB_USERNAME = 'l33ton'
GITHUB_REPO = 'Erasmus-Document-Generator'

if getattr(sys, 'frozen', False):
    base_dir = Path(sys._MEIPASS) #type: ignore
else:
    base_dir = Path(__file__).resolve().parent

if __name__ == "__main__":
    root = tk.Tk()
    app = ErasmusApp(root, base_dir)
    icon_path = base_dir / 'app_icon.ico'

    if icon_path.exists():
        root.iconbitmap(str(icon_path))

    root.after(1000, lambda: check_for_updates(root, app.status_label, CURRENT_VERSION, GITHUB_USERNAME, GITHUB_REPO))
    
    root.mainloop()