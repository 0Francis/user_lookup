import tkinter as tk
import tkinter.ttk as ttk
import os

from app.gui.styles import BG_DARK, BG_MID, FG, RED_ACCENT


class SplashScreen(tk.Toplevel):
    """
    Borderless splash window shown on startup.
    Displays an optional image from `image_path` and the app title.
    Destroys itself after `duration` milliseconds.
    """

    def __init__(self, master: tk.Tk, image_path: str = "", duration: int = 2500) -> None:
        super().__init__(master)
        self.overrideredirect(True)
        self.attributes("-topmost", True)
        self.configure(bg=BG_DARK)

        self._build(image_path)
        self._center()
        self.after(duration, self.destroy)

    def _build(self, image_path: str) -> None:
        if image_path and os.path.isfile(image_path):
            try:
                from PIL import Image, ImageTk
                img = Image.open(image_path).resize((420, 200))
                self._photo = ImageTk.PhotoImage(img)
                tk.Label(self, image=self._photo, bd=0, bg=BG_DARK).pack()
            except Exception:
                self._photo = None
                self._fallback_banner()
        else:
            self._fallback_banner()

        separator = tk.Frame(self, bg=RED_ACCENT, height=2)
        separator.pack(fill="x")

        title = tk.Label(
            self,
            text="User Lookup Extractor",
            font=("Segoe UI", 14, "bold"),
            bg=BG_DARK,
            fg=FG,
            pady=12,
        )
        title.pack()

        sub = tk.Label(
            self,
            text="Loading...",
            font=("Segoe UI", 9),
            bg=BG_DARK,
            fg=RED_ACCENT,
            pady=(0),
        )
        sub.pack(pady=(0, 14))

    def _fallback_banner(self) -> None:
        """Shown when no splash image is provided or loading fails."""
        banner = tk.Frame(self, bg=RED_ACCENT, height=6)
        banner.pack(fill="x")
        spacer = tk.Frame(self, bg=BG_DARK, height=60, width=420)
        spacer.pack()

    def _center(self) -> None:
        self.update_idletasks()
        w = self.winfo_reqwidth()
        h = self.winfo_reqheight()
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        self.geometry(f"+{(sw - w) // 2}+{(sh - h) // 2}")
