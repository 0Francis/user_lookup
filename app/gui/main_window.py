import threading
import tkinter as tk
import tkinter.ttk as ttk
from tkinter import messagebox
from typing import Callable

from app.config import AppConfig
from app.extractor import ExtractionCancelledError, UserExtractor
from app.gui.styles import BG_DARK, BG_MID, FG, FG_DIM, RED_ACCENT


class GradientProgressBar(tk.Canvas):
    """
    Horizontal progress bar that interpolates from near-black red (0%)
    to vivid bright red (100%) as the fill grows.
    """

    DARK_RED   = (100, 10, 10)
    BRIGHT_RED = (220, 30, 30)
    TROUGH     = BG_DARK

    def __init__(self, master: tk.Widget, width: int = 500, height: int = 8, **kw) -> None:
        super().__init__(master, width=width, height=height,
                         bg=self.TROUGH, highlightthickness=0, **kw)
        self._width  = width
        self._height = height
        self._value  = 0.0

    def set(self, fraction: float) -> None:
        self._value = max(0.0, min(1.0, fraction))
        self._draw()

    def reset(self) -> None:
        self._value = 0.0
        self.delete("bar")

    def _lerp(self, t: float) -> str:
        r = int(self.DARK_RED[0] + (self.BRIGHT_RED[0] - self.DARK_RED[0]) * t)
        g = int(self.DARK_RED[1] + (self.BRIGHT_RED[1] - self.DARK_RED[1]) * t)
        b = int(self.DARK_RED[2] + (self.BRIGHT_RED[2] - self.DARK_RED[2]) * t)
        return f"#{r:02x}{g:02x}{b:02x}"

    def _draw(self) -> None:
        self.delete("bar")
        filled = int(self._width * self._value)
        if filled == 0:
            return
        for x in range(filled):
            t = (x / (self._width - 1)) * self._value if self._width > 1 else self._value
            self.create_line(x, 0, x, self._height, fill=self._lerp(t), tags="bar")


class MainWindow(ttk.Frame):
    """
    Main results window: Treeview table, gradient progress bar,
    Start / Cancel / Export controls.
    """

    COLUMNS = ("Hostname", "AUUID", "Full Name")

    def __init__(
        self,
        master: tk.Tk,
        config: AppConfig,
        on_back: Callable[[], None] | None = None,
    ) -> None:
        super().__init__(master, padding=12)
        self.config   = config
        self.on_back  = on_back
        self._extractor: UserExtractor | None = None
        self._records:   list[tuple[str, str, str]] = []

        self.grid(row=0, column=0, sticky="nsew")
        master.rowconfigure(0, weight=1)
        master.columnconfigure(0, weight=1)
        self.rowconfigure(1, weight=1)
        self.columnconfigure(0, weight=1)

        self._build_toolbar()
        self._build_tree()
        self._build_progress()
        self._build_statusbar()

    def _build_toolbar(self) -> None:
        bar = ttk.Frame(self)
        bar.grid(row=0, column=0, sticky="ew", pady=(0, 8))

        self._start_btn = ttk.Button(
            bar, text="▶  Start", style="Accent.TButton", command=self.start,
        )
        self._start_btn.pack(side="left", padx=(0, 6))

        self._cancel_btn = ttk.Button(
            bar, text="✕  Cancel", command=self.cancel, state="disabled",
        )
        self._cancel_btn.pack(side="left", padx=(0, 6))

        self._export_btn = ttk.Button(
            bar, text="💾  Export", command=self.export, state="disabled",
        )
        self._export_btn.pack(side="left")

        if self.on_back:
            back_btn = ttk.Button(bar, text="← Back", command=self.on_back)
            back_btn.pack(side="right")

    def _build_tree(self) -> None:
        frame = ttk.Frame(self)
        frame.grid(row=1, column=0, sticky="nsew")
        frame.rowconfigure(0, weight=1)
        frame.columnconfigure(0, weight=1)

        self._tree = ttk.Treeview(
            frame, columns=self.COLUMNS, show="headings", height=22,
        )
        for col in self.COLUMNS:
            self._tree.heading(col, text=col)
            self._tree.column(col, minwidth=80, stretch=True)
        self._tree.column("Hostname",  width=160)
        self._tree.column("AUUID",     width=100)
        self._tree.column("Full Name", width=240)
        self._tree.grid(row=0, column=0, sticky="nsew")

        vsb = ttk.Scrollbar(frame, orient="vertical", command=self._tree.yview)
        self._tree.configure(yscrollcommand=vsb.set)
        vsb.grid(row=0, column=1, sticky="ns")

    def _build_progress(self) -> None:
        self._progress = GradientProgressBar(self, height=8)
        self._progress.grid(row=2, column=0, sticky="ew", pady=(8, 0))

    def _build_statusbar(self) -> None:
        self._status_var = tk.StringVar(value="Ready.")
        status = tk.Label(
            self, textvariable=self._status_var,
            bg=BG_DARK, fg=FG_DIM, font=("Segoe UI", 8),
            anchor="w",
        )
        status.grid(row=3, column=0, sticky="ew", pady=(4, 0))

    def start(self) -> None:
        self._tree.delete(*self._tree.get_children())
        self._records.clear()
        self._progress.reset()
        self._extractor = UserExtractor(self.config)
        self._start_btn.config(state="disabled")
        self._cancel_btn.config(state="normal")
        self._export_btn.config(state="disabled")
        self._status_var.set("Extracting…")
        threading.Thread(target=self._run_extraction, daemon=True).start()

    def cancel(self) -> None:
        if self._extractor:
            self._extractor.cancel()
        self._cancel_btn.config(state="disabled")
        self._status_var.set("Cancelling…")

    def export(self) -> None:
        if not self._records:
            messagebox.showwarning("No data", "Nothing to export yet.")
            return
        try:
            extractor = UserExtractor(self.config)
            extractor.save_results(self._records)
            messagebox.showinfo("Exported", f"Saved to:\n{self.config.output_file}")
        except Exception as e:
            messagebox.showerror("Export failed", str(e))

    def _run_extraction(self) -> None:
        try:
            records = self._extractor.extract(progress_cb=self._on_progress)
            self._records = records
            self.after(0, lambda: self._update_tree(records))
            self.after(0, lambda: self._status_var.set(
                f"Done — {len(records)} record(s) resolved."
            ))
            self.after(0, lambda: self._export_btn.config(state="normal"))
        except ExtractionCancelledError:
            self.after(0, lambda: self._status_var.set("Cancelled."))
        except Exception as e:
            self.after(0, lambda: messagebox.showerror("Extraction error", str(e)))
            self.after(0, lambda: self._status_var.set("Error — see dialog."))
        finally:
            self.after(0, lambda: self._start_btn.config(state="normal"))
            self.after(0, lambda: self._cancel_btn.config(state="disabled"))

    def _on_progress(self, done: int, total: int) -> None:
        fraction = done / total if total else 0
        self.after(0, lambda: self._progress.set(fraction))
        self.after(0, lambda: self._status_var.set(f"Resolving {done}/{total}…"))

    def _update_tree(self, records: list[tuple[str, str, str]]) -> None:
        self._tree.delete(*self._tree.get_children())
        for rec in records:
            self._tree.insert("", "end", values=rec)
