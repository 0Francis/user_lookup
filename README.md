# User Lookup Extractor

Resolves Active Directory full names from AUUID-embedded hostnames stored in an Excel file. Runs on Windows, requires domain connectivity for live lookups. Packagable as a standalone `.exe` via PyInstaller.

---

## Quick Start

```powershell
# 1. Clone / navigate to the project
cd user_lookup

# 2. Create and activate the virtual environment
python -m venv venv
venv\Scripts\activate

# 3. Install dependencies
pip install -r requirements.txt

# 4. Run
python main.py
```

---

## Application Flow

```
Splash screen (2.5 s)
  └── Setup screen - fill in 5 fields, each validates live with tick/cross
        └── Main window - Start extraction, watch gradient progress bar, Export to Excel
                          ← Back returns to Setup
```

---

## Setup Screen Fields

| Field | What to enter |
|---|---|
| **Excel file path** | Full path to the source `.xlsx` file |
| **Sheet name** | Worksheet name containing hostname data |
| **Hostname column** | Column header that holds hostnames / AUUIDs |
| **Output file (.xlsx)** | Full path for the output file (must end `.xlsx`) |
| **Rows to extract** | Positive integer, or `all` |

Fields validate in real time as you type. **Continue** stays disabled until all 5 show a green ✓. Last-used values are remembered across runs (`config/settings.json`).

---

## Project Structure

```
user_lookup/
├── main.py                  ← entry point
├── requirements.txt         ← runtime deps
├── requirements-dev.txt     ← adds pytest + pyinstaller
├── .env.example             ← optional CI/test path pre-seeds
├── assets/
│   ├── splash.webp          ← splash image
│   └── splash.ico           ← app icon
├── config/
│   └── settings.json        ← auto-created on first run
├── app/
│   ├── config.py            ← AppConfig dataclass + load/save/resource_path
│   ├── extractor.py         ← UserExtractor (pure logic, no GUI)
│   ├── gui/
│   │   ├── styles.py        ← dark theme + red accent
│   │   ├── splash.py        ← SplashScreen widget
│   │   ├── input_screen.py  ← live-validated parameter form
│   │   └── main_window.py   ← results table + GradientProgressBar
│   └── utils/
│       └── encoding.py      ← Windows console codepage decoder
└── tests/
    ├── test_config.py
    ├── test_encoding.py
    └── test_extractor.py
```

---

## Adding Assets

| File | Notes |
|---|---|
| `assets/splash.webp` | Shown on startup. Recommended size: 420x200 px. Any format Pillow supports. |
| `assets/splash.ico` | Window/taskbar icon. Standard `.ico` multi-size file. |

Both are optional. The app runs fine without them and falls back to a plain red banner.

---

## Running Tests

```powershell
venv\Scripts\activate
pytest tests\ -v
```

All 21 tests should pass. Tests cover `AppConfig` validation + JSON roundtrip, `decode_net_output` encoding fallback chain, and `UserExtractor` dataframe parsing, AD query parsing, cancellation, and Excel export.

---

## Building a Standalone .exe (PyInstaller)

```powershell
# Install dev deps (includes pyinstaller)
pip install -r requirements-dev.txt

# Build
pyinstaller user_lookup.spec --clean
```

Output: `dist/UserLookup.exe` - single portable executable, no Python required on the target machine.

---

## Dependencies

| Package | Purpose |
|---|---|
| `pandas` | Excel reading + dataframe operations |
| `openpyxl` | Excel engine |
| `Pillow` | Splash screen image loading |
| `tkinter` | GUI (Python stdlib) |

---

## Notes

- Requires Windows domain connectivity for `net user /domain` lookups. Offline machines will see `AD error: ...` in the Full Name column. The app does not crash.
- `config/settings.json` is written next to the `.exe` when frozen so it persists between runs without being bundled inside the executable.
- `user_lookup20.py` is a separate standalone version and is not part of this codebase.
