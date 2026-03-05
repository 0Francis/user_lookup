import pandas as pd
import subprocess
import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox
import re
import ctypes
import locale
import unicodedata

RAW_FILE = r"c:\Users\23225632\Downloads\Kenya Offrole & CWK Dump_27 FEB.xlsx"
SHEET_NAME = "Sheet5"
HOSTNAME_COL = "Hostname"
OUTPUT_FILE = "output.xlsx"

PRIMARY_BG = "#FFFFFF"
ACCENT_RED = "#D32F2F"
TEXT_COLOR = "#222222"
SUBTEXT_COLOR = "#555555"
FONT_FAMILY = "Segoe UI"
BASE_PAD = 10


def extract_valid_ids(df, col):
    # Defensive copy to avoid SettingWithCopy warnings
    df = df.copy()
    df = df[df[col].notnull() & (df[col].astype(str).str.strip() != "")]
    # Extract first sequence of digits with at least 4 digits, take first 8 digits only
    df[col] = (
        df[col]
        .astype(str)
        .str.extract(r"(\d{4,})", expand=False)
        .str.slice(0, 7)
    )
    # Only keep IDs with at least 4 digits
    return (
        df[df[col].notnull() & df[col].astype(str).str.match(r"^\d{4,}$")][col]
        .astype(str)
        .unique()
        .str.slice(0, 7)
    )

# ---- Unicode-safe decoding for `net user` output ----
def _decode_net_output(raw: bytes) -> str:
    # 1) Try the current console output code page (OEM)
    try:
        cp = ctypes.windll.kernel32.GetConsoleOutputCP()
        if cp:
            enc = f"cp{cp}"
            text = raw.decode(enc, errors="strict")
            return unicodedata.normalize("NFC", text)
    except Exception:
        pass

    # 2) Preferred system encoding + common fallbacks
    fallbacks = [
        locale.getpreferredencoding(False),
        "mbcs",
        "cp65001",
        "utf-8",
        "cp1252",
        "cp850",
        "cp437",
    ]
    tried = set()
    for enc in fallbacks:
        if not enc:
            continue
        enc_l = enc.lower()
        if enc_l in tried:
            continue
        tried.add(enc_l)
        try:
            text = raw.decode(enc, errors="strict")
            return unicodedata.normalize("NFC", text)
        except Exception:
            continue

    # 3) Last resort: decode with replacement (better than crashing)
    try:
        text = raw.decode("mbcs", errors="replace")
    except Exception:
        text = raw.decode(errors="replace")
    return unicodedata.normalize("NFC", text)


def get_user_info(user_id):
    try:
        # Capture BYTES and decode ourselves for Windows compatibility.
        result = subprocess.run(
            ["cmd", "/c", f"net user /domain {user_id}"],
            capture_output=True,
            check=True,
        )
        output = _decode_net_output(result.stdout)

        fullname = ""
        local_groups = []
        global_groups = []
        lines = output.splitlines()

        for i, line in enumerate(lines):
            s = line.strip()

            if s.startswith("Full Name"):
                parts = re.split(r"\s{2,}", s)
                if len(parts) > 1:
                    fullname = parts[-1].strip()

            if s.startswith("Local Group Memberships"):
                parts = re.split(r"\s{2,}", s)
                if len(parts) > 1:
                    local_groups.extend(parts[1:])
                for next_line in lines[i + 1 :]:
                    ns = next_line.strip()
                    if (
                        ns == ""
                        or ns.startswith("Global Group memberships")
                        or ns.startswith("The command completed successfully")
                    ):
                        break
                    group_names = [g for g in ns.split() if g.startswith("*")]
                    local_groups.extend(group_names)

            if s.startswith("Global Group memberships"):
                parts = re.split(r"\s{2,}", s)
                if len(parts) > 1:
                    global_groups.extend(parts[1:])
                for next_line in lines[i + 1 :]:
                    ns = next_line.strip()
                    if ns == "" or ns.startswith("The command completed successfully"):
                        break
                    group_names = [g for g in ns.split() if g.startswith("*")]
                    global_groups.extend(group_names)

        local_groups = [g.replace("*", "").strip() for g in local_groups if g.strip()]
        global_groups = [g.replace("*", "").strip() for g in global_groups if g.strip()]
        return fullname, local_groups, global_groups

    except Exception:
        # If command fails or user not found, return blank but safe values
        return "", [], []


def apply_base_style(root):
    root.configure(bg=PRIMARY_BG)

    style = ttk.Style()
    try:
        style.theme_use("clam")
    except Exception:
        pass

    style.configure("App.TFrame", background=PRIMARY_BG)
    style.configure("Card.TFrame", background=PRIMARY_BG)

    style.configure(
        "App.TLabel",
        background=PRIMARY_BG,
        foreground=TEXT_COLOR,
        font=(FONT_FAMILY, 10),
    )
    style.configure(
        "Hint.TLabel",
        background=PRIMARY_BG,
        foreground=SUBTEXT_COLOR,
        font=(FONT_FAMILY, 9),
    )
    style.configure(
        "Title.TLabel",
        background=PRIMARY_BG,
        foreground=TEXT_COLOR,
        font=(FONT_FAMILY, 12, "bold"),
    )
    style.configure(
        "App.TButton",
        font=(FONT_FAMILY, 10, "semibold"),
        padding=6,
    )
    style.map("App.TButton", background=[("active", "#efefef")])

    # Treeview
    style.configure(
        "App.Treeview",
        background="#FFFFFF",
        foreground=TEXT_COLOR,
        fieldbackground="#FFFFFF",
        rowheight=24,
        font=(FONT_FAMILY, 10),
        borderwidth=0,
    )
    style.configure(
        "App.Treeview.Heading",
        background="#F5F5F5",
        foreground=TEXT_COLOR,
        font=(FONT_FAMILY, 10, "bold"),
        borderwidth=0,
    )
    style.map("App.Treeview", background=[("selected", "#FFE5E5")])


def thin_bordered_frame(parent, padding=8):
    """
    Create a white 'card' that has a thin (1px) red border using highlight.
    No red blocks, just a crisp outline.
    """
    frame = tk.Frame(
        parent,
        bg=PRIMARY_BG,
        highlightbackground=ACCENT_RED,
        highlightcolor=ACCENT_RED,
        highlightthickness=1,
        bd=0,
    )
    inner = tk.Frame(frame, bg=PRIMARY_BG, bd=0, highlightthickness=0)
    inner.pack(padx=padding, pady=padding, fill="both", expand=True)
    return frame, inner


def thin_border_widget(widget):
    try:
        widget.configure(
            highlightbackground=ACCENT_RED,
            highlightcolor=ACCENT_RED,
            highlightthickness=1,
            bd=0,
            relief="flat",
        )
    except tk.TclError:
        # Some widgets (e.g., ttk) don't take these options
        pass


def request_cancel():
    global cancel_requested
    cancel_requested = True


def process_ids():
    global cancel_requested
    cancel_requested = False
    extract_button.config(state="disabled", text="Working…")
    cancel_button.config(state="normal")
    root.update_idletasks()

    ids = text_area.get("1.0", tk.END).strip().splitlines()
    raw_codes = country_code_var.get().strip().upper()
    selected_codes = [code.strip() for code in re.split(r",|\s", raw_codes) if code.strip()]

    results = []
    tree.delete(*tree.get_children())

    def has_code(groups, codes):
        return any(any(code in g for code in codes) for g in groups)

    for user_id in ids:
        if cancel_requested:
            break
        if not user_id.strip():
            continue

        fullname, local_groups, global_groups = get_user_info(user_id)

        if not selected_codes:
            results.append(
                {
                    "ID": user_id,
                    "Fullname": fullname,
                    "LocalGroups": ", ".join(local_groups),
                    "GlobalGroups": ", ".join(global_groups),
                }
            )
            tree.insert("", "end", values=(user_id, fullname, ", ".join(local_groups), ", ".join(global_groups)))
        else:
            if has_code(local_groups, selected_codes) or has_code(global_groups, selected_codes):
                filtered_local = [g for g in local_groups if any(code in g for code in selected_codes)]
                filtered_global = [g for g in global_groups if any(code in g for code in selected_codes)]
                results.append(
                    {
                        "ID": user_id,
                        "Fullname": fullname,
                        "LocalGroups": ", ".join(filtered_local),
                        "GlobalGroups": ", ".join(filtered_global),
                    }
                )
                tree.insert("", "end", values=(user_id, fullname, ", ".join(filtered_local), ", ".join(filtered_global)))

        root.update_idletasks()

    try:
        out_df = pd.DataFrame(results)
        out_df.to_excel(OUTPUT_FILE, index=False, engine="openpyxl")
        messagebox.showinfo("Done", f"Results saved to {OUTPUT_FILE}")
    except Exception as e:
        messagebox.showerror("Save Error", f"Could not save to {OUTPUT_FILE}\n\n{e}")

    extract_button.config(state="normal", text="Extract Fullnames")
    cancel_button.config(state="disabled")
    root.update_idletasks()


root = tk.Tk()
root.title("Domain Fullname Extractor")
root.geometry("900x600")
root.minsize(800, 520)
apply_base_style(root)

# Root padding frame
container = ttk.Frame(root, style="App.TFrame", padding=BASE_PAD)
container.pack(fill="both", expand=True)

# Title/Header (thin border)
title_card, title_inner = thin_bordered_frame(container, padding=8)
title_card.pack(fill="x", pady=(0, BASE_PAD))
ttk.Label(title_inner, text="Domain Fullname Extractor", style="Title.TLabel").pack(anchor="w")
ttk.Label(
    title_inner,
    text="Paste IDs, optionally filter by country codes (KE, RW, TZ, ZM, SC, etc.), then extract.",
    style="Hint.TLabel",
).pack(anchor="w", pady=(2, 0))

# Top controls (IDs + Country + Button) — thin border
top_card, top = thin_bordered_frame(container, padding=12)
top_card.pack(fill="x", pady=(0, BASE_PAD))

# IDs block (left)
ids_col = ttk.Frame(top, style="Card.TFrame")
ids_col.grid(row=0, column=0, sticky="nsew", padx=(0, BASE_PAD))
ttk.Label(ids_col, text="Paste IDs (one per line):", style="App.TLabel").pack(anchor="w", pady=(0, 4))

# Scrolled text with thin red border
text_area = scrolledtext.ScrolledText(
    ids_col,
    width=40,
    height=10,
    bg=PRIMARY_BG,
    fg=TEXT_COLOR,
    insertbackground=TEXT_COLOR,
    relief="flat",
    font=(FONT_FAMILY, 10),
)
thin_border_widget(text_area)
text_area.pack(fill="both", expand=True)

# Right column (Country + Button)
right_col = ttk.Frame(top, style="Card.TFrame")
right_col.grid(row=0, column=1, sticky="nsew")

ttk.Label(right_col, text="Country Code(s)", style="App.TLabel").pack(anchor="w", pady=(0, 4))
ttk.Label(
    right_col,
    text="Example: KE, RW, TZ, ZM, SC (use comma or space)",
    style="Hint.TLabel",
).pack(anchor="w", pady=(0, 6))

country_code_var = tk.StringVar()

# Entry with thin red border
entry_holder = tk.Frame(right_col, bg=PRIMARY_BG)
entry_holder.pack(fill="x", pady=(0, 10))
country_code_entry = tk.Entry(
    entry_holder,
    textvariable=country_code_var,
    bg=PRIMARY_BG,
    fg=TEXT_COLOR,
    relief="flat",
    font=(FONT_FAMILY, 10),
    insertbackground=TEXT_COLOR,
)
thin_border_widget(country_code_entry)
country_code_entry.pack(fill="x")

# Button (ttk)
extract_button = ttk.Button(right_col, text="Extract Fullnames", command=process_ids)
extract_button.pack(pady=(10, 0), anchor="center")

cancel_button = ttk.Button(right_col, text="Cancel Extraction", command=request_cancel, state="disabled")
cancel_button.pack(pady=(2, 0), anchor="center")

# Treeview section — thin border around the whole table
table_card, table_wrap = thin_bordered_frame(container, padding=10)
table_card.pack(fill="both", expand=True)

columns = ("ID", "Fullname", "LocalGroups", "GlobalGroups")

# Holder for tree + scrollbars
table_holder = tk.Frame(table_wrap, bg=PRIMARY_BG)
table_holder.pack(fill="both", expand=True)

# Create tree inside table_holder
tree = ttk.Treeview(
    table_holder,
    columns=columns,
    show="headings",
    height=10,
    style="App.Treeview",
)
for col in columns:
    tree.heading(col, text=col)
    if col == "ID":
        tree.column(col, width=120, anchor="w")
    elif col == "Fullname":
        tree.column(col, width=220, anchor="w")
    else:
        tree.column(col, width=260, anchor="w")

# Scrollbars (plain ttk; no red blocks)
yscroll = ttk.Scrollbar(table_holder, orient="vertical", command=tree.yview)
xscroll = ttk.Scrollbar(table_holder, orient="horizontal", command=tree.xview)
tree.configure(yscrollcommand=yscroll.set, xscrollcommand=xscroll.set)

# Pack: tree left, vertical scrollbar to its right, horizontal at bottom
tree.pack(side="left", fill="both", expand=True)
yscroll.pack(side="left", fill="y", padx=(6, 0))
xscroll.pack(side="bottom", fill="x", pady=(6, 0))

# Make the top grid columns responsive
top.grid_columnconfigure(0, weight=3)
top.grid_columnconfigure(1, weight=2)

root.mainloop()