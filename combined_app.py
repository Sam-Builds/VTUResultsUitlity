import sys
import os
import threading
import tkinter as tk
from tkinter import ttk
from tkinterdnd2 import TkinterDnD

from app import VTUParserApp
from scraper import initialize_ocr

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

class CombinedApp(TkinterDnD.Tk):
    def __init__(self):
        super().__init__()
        
        self.title("VTU Results Utility - Professional Edition")
        screen_w = self.winfo_screenwidth()
        screen_h = self.winfo_screenheight()
        win_w = max(720, min(920, int(screen_w * 0.50)))
        win_h = max(640, min(780, int(screen_h * 0.76)))
        pos_x = max(20, (screen_w - win_w) // 2)
        pos_y = max(20, (screen_h - win_h) // 2)
        self.geometry(f"{win_w}x{win_h}+{pos_x}+{pos_y}")
        self.configure(bg="#f5f7fa")
        self.minsize(700, 620)
        
        try:
            ico_path = resource_path("combined.ico")
            png_path = resource_path("combined.png")

            if os.path.exists(ico_path):
                try:
                    self.iconbitmap(ico_path)
                except tk.TclError:
                    pass

            if os.path.exists(png_path):
                try:
                    self._icon_img = tk.PhotoImage(file=png_path)
                    self.iconphoto(True, self._icon_img)
                except tk.TclError:
                    pass
        except Exception as e:
            print(f"Could not load icon: {e}")

        self._show_startup_splash()
        self.after(120, self._initialize_main_ui)

        self.protocol("WM_DELETE_WINDOW", self._on_close)

    def _show_startup_splash(self):
        self.withdraw()
        self.splash = tk.Toplevel(self)
        self.splash.title("Loading")
        self.splash.configure(bg="#f5f7fa")
        self.splash.overrideredirect(True)

        w, h = 380, 170
        screen_w = self.winfo_screenwidth()
        screen_h = self.winfo_screenheight()
        x = (screen_w - w) // 2
        y = (screen_h - h) // 2
        self.splash.geometry(f"{w}x{h}+{x}+{y}")

        tk.Label(
            self.splash,
            text="VTU Results Utility",
            font=("Segoe UI", 14, "bold"),
            bg="#f5f7fa",
            fg="#1a73e8",
        ).pack(pady=(28, 12))

        tk.Label(
            self.splash,
            text="Initializing parser and OCR engine...",
            font=("Segoe UI", 10),
            bg="#f5f7fa",
            fg="#5f6368",
        ).pack(pady=(0, 12))

        bar = ttk.Progressbar(self.splash, mode="indeterminate", length=260)
        bar.pack(pady=(0, 12))
        bar.start(12)

    def _initialize_main_ui(self):
        threading.Thread(target=initialize_ocr, daemon=True).start()

        self.parser_app = VTUParserApp(self)
        self.parser_app.pack(fill="both", expand=True)

        if hasattr(self, "splash") and self.splash.winfo_exists():
            self.splash.destroy()
        self.deiconify()

    def _on_close(self):
        if hasattr(self, "parser_app") and not self.parser_app.can_close():
            return
        self.destroy()

if __name__ == "__main__":
    app = CombinedApp()
    app.mainloop()
