# Proposed Improvements — User Lookup Extractor

A structured breakdown of what to improve, why, and how to make this PyInstaller-ready.

---

## 1. Project Structure (Proposed)

Reorganise from a flat pile of scripts into a proper package:

```
user_lookup/
├── main.py                   ← single entry point
├── user_lookup.spec          ← PyInstaller build spec
├── requirements.txt
├── assets/
│   ├── splash.png            ← splash screen image
│   └── icon.ico              ← app icon
├── config/
│   └── settings.json         ← persisted last-used config (auto-created on first run)
└── app/
    ├── __init__.py
    ├── config.py             ← Config dataclass + load/save logic
    ├── extractor.py          ← pure data/domain logic (no GUI)
    ├── gui/
    │   ├── __init__.py
    │   ├── splash.py         ← SplashScreen widget
    │   ├── input_screen.py   ← InputScreen widget
    │   ├── main_window.py    ← UserLookupGUI (main table + controls)
    │   └── styles.py         ← ttk.Style definitions (red gradient border theme)
    └── utils/
        ├── __init__.py
        └── encoding.py       ← _decode_net_output (isolated, testable)
```

**Why:** Right now all logic (data reading, AD querying, GUI, config) is mixed together. Splitting by responsibility means each piece can be changed, tested, or swapped without breaking the others.

---

## 2. Proper OOP — Class Responsibilities

### Current problems
- `UserLookupGUI` does everything: reads Excel, runs `net user`, parses output, updates widgets.
- Config is module-level globals injected from outside — fragile.
- The launcher and core are tightly coupled (`import user_lookup12 as ul; ul.RAW_FILE = ...`).

### Proposed class design

```
AppConfig (dataclass)
    fields: raw_file, sheet_name, hostname_col, output_file, num_rows
    methods: load_from_json(), save_to_json(), validate()

UserExtractor
    __init__(self, config: AppConfig)
    extract() → generator/list of (hostname, auuid, full_name)
    get_user_info(auuid: str) → str
    _decode_net_output(raw: bytes) → str   (or delegate to utils.encoding)

App  (orchestrator — owns root Tk window)
    __init__()
    run()                 ← mainloop
    _show_splash()
    _show_input()
    _show_main(config)

SplashScreen(tk.Toplevel)
    __init__(master, image_path, duration)

InputScreen(ttk.Frame)
    __init__(master, config: AppConfig, on_submit)
    _on_field_change(key, var)         ← bound to each entry via StringVar trace
    _validate_field(key, value) → bool ← per-field async validation
    _set_tick(key, state)              ← shows ✓ / ✗ / spinner next to the field
    _all_valid() → bool
    submit()

MainWindow(ttk.Frame)
    __init__(master, config: AppConfig, extractor: UserExtractor)
    start()
    cancel()
    export()
    _update_tree(records)
```

**Key principle:** GUI classes only handle display and user events. All data work stays in `UserExtractor`. `App` wires them together.

---

## 3. Execution Flow (Proposed)

```
main.py
  └── App.__init__()
        └── _show_splash(assets/splash.png, 2500ms)
              └── _show_input()  ← pre-filled from last saved config
                    └── on_submit(config)
                          ├── config.save_to_json()          ← remember last inputs
                          └── _show_main(config)
                                └── MainWindow
                                      ├── Start → UserExtractor.extract() [thread]
                                      │     └── progress callback → update progress bar
                                      ├── Cancel → extractor.cancel()
                                      └── Export → write to Excel
```

**Improvements over current:**
- Config is saved to `config/settings.json` so the user doesn't re-type paths every run.
- Each input field validates live as the user types — a tick appears when valid.
- Progress bar replaces the silent wait during extraction, animating dark → bright red.
- Cancellation is a proper flag on `UserExtractor`, not a GUI attribute.

---

## 4. Input Field Validation with Tick Indicators

Each field in `InputScreen` validates itself as soon as the user stops typing (debounced 400 ms). A small label next to each entry shows the status.

### Validation rules per field

| Field | Validation logic |
|---|---|
| `raw_file` | File exists **and** `pd.read_excel(path, nrows=1)` succeeds |
| `sheet_name` | Checked against `pd.ExcelFile(raw_file).sheet_names` (after raw_file is valid) |
| `hostname_col` | Checked against `df.columns` from the first read |
| `output_file` | Parent directory exists and name ends with `.xlsx` |
| `num_rows` | Is a positive integer, or the literal string `all` |

### How it works

```python
# app/gui/input_screen.py
import threading, time

class InputScreen(ttk.Frame):
    DEBOUNCE_MS = 400

    def __init__(self, master, config, on_submit):
        super().__init__(master, padding=20)
        self._vars   = {}   # key → tk.StringVar
        self._ticks  = {}   # key → ttk.Label  (the ✓ / ✗ indicator)
        self._timers = {}   # key → after-id   (debounce)
        self._valid  = {}   # key → bool
        self._excel_cache = {}  # raw_file path → (sheet_names, columns)

        fields = [
            ("raw_file",      "Excel file path"),
            ("sheet_name",    "Sheet name"),
            ("hostname_col",  "Hostname column"),
            ("output_file",   "Output file (.xlsx)"),
            ("num_rows",      "Rows to extract (number or 'all')"),
        ]
        for i, (key, label) in enumerate(fields):
            ttk.Label(self, text=label).grid(row=i, column=0, sticky="w", pady=6)
            var = tk.StringVar(value=config.__dict__.get(key, ""))
            entry = ttk.Entry(self, textvariable=var, width=40)
            entry.grid(row=i, column=1, pady=6, sticky="ew")
            tick = ttk.Label(self, text="", width=2, font=("Segoe UI", 11))
            tick.grid(row=i, column=2, padx=(4, 0))
            self._vars[key]  = var
            self._ticks[key] = tick
            self._valid[key] = False
            var.trace_add("write", lambda *_, k=key: self._schedule_validate(k))

        self._submit_btn = ttk.Button(self, text="Continue",
                                      command=self.submit, state="disabled")
        self._submit_btn.grid(row=len(fields), column=0, columnspan=3, pady=15)
        self.grid_columnconfigure(1, weight=1)

    def _schedule_validate(self, key):
        """Debounce: cancel pending timer and restart."""
        if key in self._timers:
            self.after_cancel(self._timers[key])
        self._timers[key] = self.after(self.DEBOUNCE_MS,
                                       lambda: self._run_validate(key))

    def _run_validate(self, key):
        """Run validation on a background thread so the GUI doesn't freeze."""
        self._set_tick(key, "wait")
        threading.Thread(target=self._validate_field, args=(key,),
                         daemon=True).start()

    def _validate_field(self, key):
        value = self._vars[key].get().strip()
        ok = False
        try:
            if key == "raw_file":
                pd.read_excel(value, nrows=1, engine="openpyxl")
                self._excel_cache["path"] = value
                ef = pd.ExcelFile(value, engine="openpyxl")
                self._excel_cache["sheets"]  = ef.sheet_names
                ok = True
            elif key == "sheet_name":
                ok = value in self._excel_cache.get("sheets", [])
                if ok:
                    df = pd.read_excel(
                        self._excel_cache["path"], sheet_name=value,
                        nrows=1, engine="openpyxl"
                    )
                    self._excel_cache["columns"] = list(df.columns)
            elif key == "hostname_col":
                ok = value in self._excel_cache.get("columns", [])
            elif key == "output_file":
                import os
                ok = (value.endswith(".xlsx") and
                      os.path.isdir(os.path.dirname(os.path.abspath(value))))
            elif key == "num_rows":
                ok = (value.lower() == "all" or
                      (value.isdigit() and int(value) > 0))
        except Exception:
            ok = False
        self._valid[key] = ok
        self.after(0, lambda: self._set_tick(key, "ok" if ok else "fail"))
        self.after(0, self._refresh_submit)

    def _set_tick(self, key, state):
        """Update the indicator label. Must be called from the main thread."""
        symbols = {"ok": ("✓", "#2ECC71"), "fail": ("✗", "#E74C3C"),
                   "wait": ("…", "#AAAAAA")}
        text, colour = symbols[state]
        self._ticks[key].config(text=text, foreground=colour)

    def _refresh_submit(self):
        state = "normal" if all(self._valid.values()) else "disabled"
        self._submit_btn.config(state=state)

    def submit(self):
        if not all(self._valid.values()):
            return
        values = {k: v.get().strip() for k, v in self._vars.items()}
        self.on_submit(values)
```

**Result:** The **Continue** button stays disabled until every field shows a green ✓. Sheet name and hostname column fields automatically validate against data read from the already-confirmed Excel file, so the user gets errors at input time, not mid-extraction.

---

## 5. Splash Screen with Image

```python
# app/gui/splash.py
from PIL import Image, ImageTk   # requires Pillow

class SplashScreen(tk.Toplevel):
    def __init__(self, master, image_path: str, duration: int = 2500):
        super().__init__(master)
        self.overrideredirect(True)          # no title bar
        self.attributes("-topmost", True)

        img = Image.open(image_path).resize((400, 220))
        self._photo = ImageTk.PhotoImage(img)

        tk.Label(self, image=self._photo, bd=0).pack()
        ttk.Label(self, text="User Lookup Extractor", 
                  font=("Segoe UI", 13, "bold")).pack(pady=(0, 10))

        self._center()
        self.after(duration, self.destroy)

    def _center(self):
        self.update_idletasks()
        w, h = self.winfo_width(), self.winfo_height()
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        self.geometry(f"+{(sw-w)//2}+{(sh-h)//2}")
```

**Dependency to add:** `Pillow` — needed for image loading.

---

## 5. Red Gradient Border Styling

Tkinter's `ttk` doesn't support true CSS-style gradients, but you can achieve a clean thin red accent border with a custom theme:

```python
# app/gui/styles.py
import tkinter.ttk as ttk

RED_ACCENT   = "#C0392B"    # solid red for borders/highlights
RED_LIGHT    = "#E74C3C"    # lighter red for hover states
BG_DARK      = "#1E1E1E"
BG_MID       = "#2D2D2D"
FG           = "#F0F0F0"

def apply_theme(root):
    style = ttk.Style(root)
    style.theme_use("clam")

    # Main frame — thin red left border feel via a colored separator
    style.configure("TFrame",      background=BG_MID)
    style.configure("TLabel",      background=BG_MID, foreground=FG)
    style.configure("TEntry",      fieldbackground=BG_DARK, foreground=FG,
                                   insertcolor=FG)

    # Buttons with red accent
    style.configure("Accent.TButton",
                    background=RED_ACCENT, foreground="white",
                    font=("Segoe UI", 9, "bold"), relief="flat", padding=6)
    style.map("Accent.TButton",
              background=[("active", RED_LIGHT), ("disabled", "#555")])

    # Treeview — red heading bar, dark rows
    style.configure("Treeview.Heading",
                    background=RED_ACCENT, foreground="white",
                    font=("Segoe UI", 9, "bold"), relief="flat")
    style.configure("Treeview",
                    background=BG_DARK, foreground=FG,
                    fieldbackground=BG_DARK, rowheight=24)
    style.map("Treeview", background=[("selected", RED_LIGHT)])

    # Progress bar — animated dark-to-bright red (see Section 5a below)
    style.configure("red.Horizontal.TProgressbar",
                    troughcolor=BG_DARK, background=RED_ACCENT, thickness=6)
```

For a **true thin gradient border** effect on the window edge, wrap the main frame in a `tk.Frame` painted with a canvas gradient strip (1–3px wide) on the left/top side. This is purely visual but very effective.

### 5a. Animated Dark → Bright Red Progress Bar

`ttk.Progressbar` fills with a single flat colour. To simulate a dark-to-bright gradient that grows with progress, overlay a `tk.Canvas` on top of the trough and redraw a colour-interpolated rectangle on each step.

```python
# app/gui/main_window.py

class GradientProgressBar(tk.Canvas):
    """
    A Canvas that draws a horizontal bar whose fill interpolates
    from DARK_RED (0 %) to BRIGHT_RED (100 %) as value increases.
    """
    DARK_RED   = (100, 10, 10)    # RGB at 0 %
    BRIGHT_RED = (220, 30, 30)    # RGB at 100 %
    TROUGH     = "#1E1E1E"

    def __init__(self, master, width=400, height=8, **kw):
        super().__init__(master, width=width, height=height,
                         bg=self.TROUGH, highlightthickness=0, **kw)
        self._width  = width
        self._height = height
        self._value  = 0.0        # 0.0 – 1.0

    def set(self, fraction: float):
        """fraction: 0.0 (empty) to 1.0 (full)."""
        self._value = max(0.0, min(1.0, fraction))
        self._draw()

    def _lerp_color(self, t):
        r = int(self.DARK_RED[0] + (self.BRIGHT_RED[0] - self.DARK_RED[0]) * t)
        g = int(self.DARK_RED[1] + (self.BRIGHT_RED[1] - self.DARK_RED[1]) * t)
        b = int(self.DARK_RED[2] + (self.BRIGHT_RED[2] - self.DARK_RED[2]) * t)
        return f"#{r:02x}{g:02x}{b:02x}"

    def _draw(self):
        self.delete("bar")
        filled_px = int(self._width * self._value)
        if filled_px == 0:
            return
        # Draw N vertical slices, each slightly brighter than the last
        slices = max(filled_px, 1)
        for i in range(slices):
            t = i / (self._width - 1) if self._width > 1 else 1.0
            colour = self._lerp_color(t * self._value)   # scale to current fill
            x = i
            self.create_line(x, 0, x, self._height,
                             fill=colour, tags="bar")
```

**Usage in `MainWindow`:**
```python
self._progress = GradientProgressBar(self.frame, width=500, height=8)
self._progress.grid(row=3, column=0, columnspan=2, sticky="ew", pady=(4, 0))

# Inside the extraction loop callback:
def _on_progress(self, done: int, total: int):
    self._progress.set(done / total)
    self.master.update_idletasks()
```

The bar starts as near-black red and arrives at a vivid `#DC1E1E` when extraction is complete — no external animation library needed.

---

## 6. PyInstaller — Making It Run Anywhere

### What needs to change first
| Issue | Fix |
|---|---|
| `net user /domain` requires domain connectivity | Add a graceful offline error message instead of a crash |
| Hardcoded/injected config paths | Move to `AppConfig` + JSON, resolved relative to executable |
| `assets/` folder needed at runtime | Bundle via PyInstaller `--add-data` |
| `openpyxl` engines & hidden imports | Declare in `.spec` |

### Resolving paths at runtime (critical for PyInstaller)
```python
# app/config.py
import sys, os

def resource_path(relative: str) -> str:
    """Works both in development and when frozen by PyInstaller."""
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, relative)
```

Use `resource_path("assets/splash.png")` everywhere instead of bare relative paths.

### PyInstaller spec (user_lookup.spec)
```python
# user_lookup.spec
from PyInstaller.utils.hooks import collect_data_files

a = Analysis(
    ["main.py"],
    pathex=[],
    binaries=[],
    datas=[
        ("assets/splash.png", "assets"),
        ("assets/icon.ico",   "assets"),
    ] + collect_data_files("openpyxl"),
    hiddenimports=[
        "openpyxl",
        "openpyxl.cell._writer",
        "pandas",
        "PIL",
        "PIL.Image",
        "PIL.ImageTk",
    ],
    hookspath=[],
    noarchive=False,
)
pyz = PYZ(a.pure)
exe = EXE(
    pyz, a.scripts, a.binaries, a.zipfiles, a.datas,
    name="UserLookup",
    debug=False,
    console=False,          # no console window
    icon="assets/icon.ico",
)
```

### Build command
```
pyinstaller user_lookup.spec --clean
```

Output: `dist/UserLookup.exe` — single portable executable.

---

## 7. What to Implement First (Priority Order)

| # | Task | Effort | Impact |
|---|---|---|---|
| 1 | Refactor into `app/` package + `main.py` entry point | Medium | Unlocks everything else |
| 2 | `AppConfig` dataclass with JSON save/load | Low | User no longer re-types paths |
| 3 | Move extraction logic to `UserExtractor` | Medium | Testable, decoupled |
| 4 | `apply_theme()` + red accent styling | Low | Immediate visual lift |
| 5 | Splash screen with image (Pillow) | Low | Polished feel |
| 6 | Field validation with live ✓/✗ ticks per input | Low | Catches bad paths before extraction |
| 6a | Dark→bright red `GradientProgressBar` during extraction | Low | Better UX |
| 7 | `resource_path()` helper everywhere | Low | Required for PyInstaller |
| 8 | PyInstaller `.spec` + test frozen build | Medium | Final portability goal |

---

## 8. Additional Dependencies (requirements.txt additions)

```
pandas>=2.0
openpyxl>=3.1
Pillow>=10.0          # splash screen image
pyinstaller>=6.0      # build tool (dev only)
```

---

## Notes

- `user_lookup20.py` is a separate version and is **not part of this refactor path**.
- The `data/` folder and any `.xlsx` files should **not** be bundled into the `.exe` — the user provides those paths at runtime via the input form.
- Keep `config/settings.json` written next to the `.exe` (not inside `_MEIPASS`) so it persists between runs.
