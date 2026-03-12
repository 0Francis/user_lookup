import tkinter as tk

from app.config import AppConfig, resource_path
from app.gui.splash import SplashScreen
from app.gui.input_screen import InputScreen
from app.gui.main_window import MainWindow
from app.gui.styles import apply_theme, BG_DARK


class App:
    """
    Orchestrates the full application flow:
      1. Show splash screen
      2. Show input screen (pre-filled from last saved config)
      3. On submit → show main window
    """

    TITLE   = "User Lookup Extractor"
    WIDTH   = 680
    HEIGHT  = 520

    def __init__(self) -> None:
        self.root = tk.Tk()
        self.root.title(self.TITLE)
        self.root.geometry(f"{self.WIDTH}x{self.HEIGHT}")
        self.root.minsize(560, 420)
        self.root.configure(bg=BG_DARK)
        self.root.rowconfigure(0, weight=1)
        self.root.columnconfigure(0, weight=1)

        apply_theme(self.root)
        self._set_icon()

        self._current_frame: tk.Widget | None = None
        self._show_splash()

    def _set_icon(self) -> None:
        try:
            icon = resource_path("assets/splash.ico")
            self.root.iconbitmap(icon)
        except Exception:
            pass

    def _show_splash(self) -> None:
        self.root.withdraw()
        image_path = resource_path("assets/splash.webp")
        splash = SplashScreen(self.root, image_path=image_path, duration=2500)
        splash.grab_set()
        self.root.after(2500, self._after_splash)

    def _after_splash(self) -> None:
        self.root.deiconify()
        self._show_input()

    def _show_input(self) -> None:
        self._clear_frame()
        config = AppConfig.load()
        self.root.title(f"{self.TITLE} Setup")
        self.root.geometry("620x380")
        self._current_frame = InputScreen(self.root, config, on_submit=self._on_input_submit)

    def _on_input_submit(self, config: AppConfig) -> None:
        self._show_main(config)

    def _show_main(self, config: AppConfig) -> None:
        self._clear_frame()
        self.root.title(self.TITLE)
        self.root.geometry(f"{self.WIDTH}x{self.HEIGHT}")
        self._current_frame = MainWindow(
            self.root, config, on_back=self._show_input,
        )

    def _clear_frame(self) -> None:
        if self._current_frame is not None:
            self._current_frame.destroy()
            self._current_frame = None

    def run(self) -> None:
        self.root.mainloop()


if __name__ == "__main__":
    App().run()
