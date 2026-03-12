import os
import threading
import tkinter as tk
import tkinter.ttk as ttk
from typing import Callable

import pandas as pd

from app.config import AppConfig
from app.gui.styles import BG_DARK, BG_MID, FG, FG_DIM, RED_ACCENT


class InputScreen(ttk.Frame):
    """
    Parameter entry form shown after the splash screen.
    Each field validates live (debounced) and shows a ✓/✗ indicator.
    The Continue button is disabled until every field is valid.
    """

    DEBOUNCE_MS = 450
    FIELDS = [
        ("raw_file",     "Excel file path"),
        ("sheet_name",   "Sheet name"),
        ("hostname_col", "Hostname column"),
        ("output_file",  "Output file (.xlsx)"),
        ("num_rows",     "Rows to extract (number or 'all')"),
    ]

    def __init__(
        self,
        master: tk.Tk,
        config: AppConfig,
        on_submit: Callable[[AppConfig], None],
    ) -> None:
        super().__init__(master, padding=24)
        self.on_submit = on_submit
        self._vars:         dict[str, tk.StringVar] = {}
        self._ticks:        dict[str, tk.Label]     = {}
        self._timers:       dict[str, str]           = {}
        self._valid:        dict[str, bool]          = {}
        self._excel_cache:  dict                     = {}

        self.grid(row=0, column=0, sticky="nsew")
        master.rowconfigure(0, weight=1)
        master.columnconfigure(0, weight=1)
        self.columnconfigure(1, weight=1)

        for i, (key, label) in enumerate(self.FIELDS):
            ttk.Label(self, text=label).grid(row=i, column=0, sticky="w", pady=7, padx=(0, 10))
            var = tk.StringVar(value=getattr(config, key, ""))
            entry = ttk.Entry(self, textvariable=var, width=44)
            entry.grid(row=i, column=1, sticky="ew", pady=7)
            tick = tk.Label(self, text="", width=2, font=("Segoe UI", 12),
                            bg=BG_MID, fg=FG)
            tick.grid(row=i, column=2, padx=(6, 0))
            self._vars[key]  = var
            self._ticks[key] = tick
            self._valid[key] = False
            var.trace_add("write", lambda *_, k=key: self._schedule(k))

        self.after(120, self._revalidate_prefilled)

        sep = ttk.Separator(self, orient="horizontal")
        sep.grid(row=len(self.FIELDS), column=0, columnspan=3, sticky="ew", pady=(14, 0))

        self._submit_btn = ttk.Button(
            self, text="Continue", style="Accent.TButton",
            command=self._submit, state="disabled",
        )
        self._submit_btn.grid(
            row=len(self.FIELDS) + 1, column=0, columnspan=3, pady=(12, 0), sticky="e",
        )

    def _revalidate_prefilled(self) -> None:
        for key, var in self._vars.items():
            if var.get().strip():
                self._run(key)

    def _schedule(self, key: str) -> None:
        if key in self._timers:
            self.after_cancel(self._timers[key])
        self._timers[key] = self.after(self.DEBOUNCE_MS, lambda: self._run(key))

    def _run(self, key: str) -> None:
        self._set_tick(key, "wait")
        threading.Thread(target=self._validate, args=(key,), daemon=True).start()

    def _validate(self, key: str) -> None:
        value = self._vars[key].get().strip()
        ok = False
        try:
            if key == "raw_file":
                pd.read_excel(value, nrows=1, engine="openpyxl")
                ef = pd.ExcelFile(value, engine="openpyxl")
                self._excel_cache = {
                    "path":    value,
                    "sheets":  ef.sheet_names,
                    "columns": [],
                }
                ok = True

            elif key == "sheet_name":
                ok = value in self._excel_cache.get("sheets", [])
                if ok:
                    df = pd.read_excel(
                        self._excel_cache["path"],
                        sheet_name=value,
                        nrows=1,
                        engine="openpyxl",
                    )
                    self._excel_cache["columns"] = list(df.columns)

            elif key == "hostname_col":
                ok = value in self._excel_cache.get("columns", [])

            elif key == "output_file":
                ok = (
                    value.endswith(".xlsx")
                    and os.path.isdir(os.path.dirname(os.path.abspath(value)))
                )

            elif key == "num_rows":
                ok = value.lower() == "all" or (value.isdigit() and int(value) > 0)

        except Exception:
            ok = False

        self._valid[key] = ok
        self.after(0, lambda: self._set_tick(key, "ok" if ok else "fail"))
        self.after(0, self._refresh_btn)

    def _set_tick(self, key: str, state: str) -> None:
        symbols = {
            "ok":   ("✓", "#2ECC71"),
            "fail": ("✗", RED_ACCENT),
            "wait": ("…", FG_DIM),
        }
        text, colour = symbols[state]
        self._ticks[key].config(text=text, fg=colour)

    def _refresh_btn(self) -> None:
        state = "normal" if all(self._valid.values()) else "disabled"
        self._submit_btn.config(state=state)

    def _submit(self) -> None:
        if not all(self._valid.values()):
            return
        config = AppConfig(
            raw_file=self._vars["raw_file"].get().strip(),
            sheet_name=self._vars["sheet_name"].get().strip(),
            hostname_col=self._vars["hostname_col"].get().strip(),
            output_file=self._vars["output_file"].get().strip(),
            num_rows=self._vars["num_rows"].get().strip(),
        )
        config.save()
        self.on_submit(config)
