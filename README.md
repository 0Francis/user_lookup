# User Lookup Extractor

A Windows domain user lookup tool that resolves Active Directory full names from hostnames stored in an Excel file.

---

## What It Does

1. Reads an Excel file containing a column of hostnames (which embed user IDs / AUUIDs).
2. Extracts numeric AUUIDs from each hostname using regex (`\d{5,}`).
3. Runs `net user /domain <AUUID>` against the Windows domain for each ID.
4. Parses the **Full Name** field from the command output.
5. Displays results (Hostname, AUUID, Full Name) in a Tkinter table.
6. Exports the resolved data to an output Excel file.

---

## Project Files

| File | Role |
|---|---|
| `user_lookup_core.ipynb` | Development notebook — prototype with hardcoded config paths |
| `user_lookup12.py` | Core app (v1.2) — `UserLookupGUI` class with blank config, injected at runtime |
| `user_lookup_gui_launcher.py` | Entry point — splash screen + input form, injects config into `user_lookup12` |
| `user_lookup20.py` | Standalone newer version (v2.0) — separate, independent |
| `data/` | Source data folder (Excel dump) |

---

## How to Run

### Via Launcher (recommended)
```
python user_lookup_gui_launcher.py
```
A splash screen appears, followed by a parameter input form:

| Field | Description |
|---|---|
| Excel file path | Full path to the source `.xlsx` file |
| Sheet name | The worksheet containing hostname data |
| Hostname column | Column name that holds hostnames/AUUIDs |
| Output file name | Path for the output `.xlsx` file |
| Rows to extract | Integer count or `all` |

### Direct (v1.2)
Edit the `CONFIG` block at the top of `user_lookup12.py` and run:
```
python user_lookup12.py
```

---

## Architecture

```
user_lookup_gui_launcher.py
    └── SplashScreen (Tkinter Toplevel, 2s)
    └── InputScreen  (collects 5 config fields)
            └── injects config → user_lookup12.UserLookupGUI
                    ├── Start Extraction  (threaded, net user /domain per AUUID)
                    ├── Cancel            (sets cancel flag mid-loop)
                    └── Load Data         (writes Treeview contents to Excel)
```

---

## Key Technical Details

- **AUUID extraction:** `r"(\d{5,})"` — matches 5+ digit sequences from hostname strings.
- **Multi-value hostnames:** Cells with newline-separated hostnames are exploded into individual rows before processing.
- **Encoding:** `_decode_net_output` handles Windows console codepage detection with graceful fallback chain (`cp<N>` → `mbcs` → `utf-8` → `cp1252` → `cp850` → `cp437`).
- **Threading:** Extraction runs on a daemon thread to keep the GUI responsive.
- **Excel I/O:** `pandas` + `openpyxl`; output appends a new sheet (`user data`) to an existing file.

---

## Dependencies

```
pandas
openpyxl
tkinter  (stdlib)
```

---

## Data

`data/Kenya Offrole & CWK Dump_27 FEB.xlsx` — source dump of off-role and contract worker hostnames used as input.
