import pandas as pd
import subprocess
import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox
import os
import re

# =========================
# CONFIG
# =========================
RAW_FILE = r"c:\Users\23225632\Downloads\Kenya Offrole & CWK Dump_27 FEB.xlsx"
SHEET_NAME = "Sheet5"
HOSTNAME_COL = "Hostname"
OUTPUT_FILE = "output.xlsx"

PRIMARY_BG = "#FFFFFF"  # White background
ACCENT_RED = "#D32F2F"  # Red for border strips
TEXT_COLOR = "#222222"
SUBTEXT_COLOR = "#555555"
FONT_FAMILY = "Segoe UI"  # Looks good on Windows
BASE_PAD = 10


# =========================
# DATA HELPERS
# =========================
def extract_valid_ids(df, col):
    # Defensive copy to avoid SettingWithCopy warnings
    df = df.copy()
    df = df[df[col].notnull() & (df[col].astype(str).str.strip() != "")]
    df[col] = df[col].astype(str).str.extract(r"(\d+)", expand=False)
    return (
        df[df[col].notnull() & df[col].astype(str).str.match(r"^\d{6,}$")][col]
        .astype(str)
        .unique()
    )


def get_user_info(user_id):
    """
    Query AD via 'net user /domain {user_id}' and parse:
    - Full Name
    - Local Group Memberships
    - Global Group memberships
    Gracefully handles odd encodings and missing fields.
    """
    try:
        result = subprocess.run(
            ["net", "user", "/domain", user_id],
            capture_output=True,
            text=True,
            check=True,
            encoding="utf-8",
            errors="replace",
        )
        output = result.stdout
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


# =========================
# UI HELPERS (STYLING)
# =========================
def apply_base_style(root):
    root.configure(bg=PRIMARY_BG)

    style = ttk.Style()
    # Use a platform-appropriate theme and then override
    try:
        style.theme_use("clam")
    except Exception:
        pass

    style.configure(
        "App.TFrame",
        background=PRIMARY_BG,
    )
    style.configure(
        "Card.TFrame",
        background=PRIMARY_BG,
    )
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
    style.map(
        "App.TButton",
        background=[("active", "#efefef")],
    )

    # Treeview styles (headers + rows)
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
    style.map("App.Treeview", background=[("selected", "#FFE5E5")])  # subtle red hue


def bordered_frame(parent, padding=8):
    """
    Create a white 'card' with a thin red strip/border around it:
    - Outer frame: red (thin)
    - Inner frame: white, where content lives
    """
    outer = tk.Frame(parent, bg=ACCENT_RED, highlightthickness=0, bd=0)
    inner = tk.Frame(outer, bg=PRIMARY_BG, highlightthickness=0, bd=0)
    inner.pack(padx=padding, pady=padding, fill="both", expand=True)
    return outer, inner


# =========================
# APP LOGIC
# =========================
def process_ids():
    extract_button.config(state="disabled", text="Working…")
    root.update_idletasks()

    ids = text_area.get("1.0", tk.END).strip().splitlines()
    # Accept multiple country codes separated by comma or space
    raw_codes = country_code_var.get().strip().upper()
    selected_codes = [code.strip() for code in re.split(r",|\s", raw_codes) if code.strip()]

    results = []
    tree.delete(*tree.get_children())

    def has_code(groups, codes):
        return any(any(code in g for code in codes) for g in groups)

    for user_id in ids:
        if not user_id.strip():
            continue

        fullname, local_groups, global_groups = get_user_info(user_id)

        if not selected_codes:
            # No country code input, show all
            results.append(
                {
                    "ID": user_id,
                    "Fullname": fullname,
                    "LocalGroups": ", ".join(local_groups),
                    "GlobalGroups": ", ".join(global_groups),
                }
            )
            tree.insert(
                "",
                "end",
                values=(user_id, fullname, ", ".join(local_groups), ", ".join(global_groups)),
            )
        else:
            # Filter based on any selected country code
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
                tree.insert(
                    "",
                    "end",
                    values=(user_id, fullname, ", ".join(filtered_local), ", ".join(filtered_global)),
                )

        root.update_idletasks()

    # Save results
    try:
        out_df = pd.DataFrame(results)
        out_df.to_excel(OUTPUT_FILE, index=False, engine="openpyxl")
        messagebox.showinfo("Done", f"Results saved to {OUTPUT_FILE}")
    except Exception as e:
        messagebox.showerror("Save Error", f"Could not save to {OUTPUT_FILE}\n\n{e}")

    extract_button.config(state="normal", text="Extract Fullnames")


# =========================
# UI CONSTRUCTION
# =========================
root = tk.Tk()
root.title("Domain Fullname Extractor")
root.geometry("900x600")
root.minsize(800, 520)
apply_base_style(root)

# Root padding frame
container = ttk.Frame(root, style="App.TFrame", padding=BASE_PAD)
container.pack(fill="both", expand=True)

# Title/Header
title_bar_outer, title_bar = bordered_frame(container, padding=8)
title_bar_outer.pack(fill="x", pady=(0, BASE_PAD))
ttk.Label(title_bar, text="Domain Fullname Extractor", style="Title.TLabel").pack(anchor="w")
ttk.Label(
    title_bar,
    text="Paste IDs, optionally filter by country codes (KE, RW, TZ, ZM, SC, etc.), then extract.",
    style="Hint.TLabel",
).pack(anchor="w", pady=(2, 0))

# Top controls (IDs + Country + Button)
top_outer, top = bordered_frame(container, padding=12)
top_outer.pack(fill="x", pady=(0, BASE_PAD))

# IDs block
ids_col = ttk.Frame(top, style="Card.TFrame")
ids_col.grid(row=0, column=0, sticky="nsew", padx=(0, BASE_PAD))
ttk.Label(ids_col, text="Paste IDs (one per line):", style="App.TLabel").pack(anchor="w", pady=(0, 4))

# Scrolled text with white bg, subtle border
text_area_frame = tk.Frame(ids_col, bg=ACCENT_RED)
text_area_frame.pack(fill="both", expand=True)
text_area_inner = tk.Frame(text_area_frame, bg=PRIMARY_BG)
text_area_inner.pack(padx=2, pady=2, fill="both", expand=True)

text_area = scrolledtext.ScrolledText(
    text_area_inner,
    width=40,
    height=10,
    bg=PRIMARY_BG,
    fg=TEXT_COLOR,
    insertbackground=TEXT_COLOR,
    relief="flat",
    font=(FONT_FAMILY, 10),
)
text_area.pack(fill="both", expand=True)

# Country codes + button
right_col = ttk.Frame(top, style="Card.TFrame")
right_col.grid(row=0, column=1, sticky="nsew")
ttk.Label(
    right_col,
    text="Country Code(s)",
    style="App.TLabel",
).pack(anchor="w", pady=(0, 4))
ttk.Label(
    right_col,
    text="Example: KE, RW, TZ, ZM, SC (use comma or space)",
    style="Hint.TLabel",
).pack(anchor="w", pady=(0, 6))

country_code_var = tk.StringVar()
entry_outer = tk.Frame(right_col, bg=ACCENT_RED)
entry_outer.pack(fill="x", pady=(0, 10))
entry_inner = tk.Frame(entry_outer, bg=PRIMARY_BG)
entry_inner.pack(padx=2, pady=2, fill="x")

country_code_entry = tk.Entry(
    entry_inner,
    textvariable=country_code_var,
    bg=PRIMARY_BG,
    fg=TEXT_COLOR,
    relief="flat",
    font=(FONT_FAMILY, 10),
    insertbackground=TEXT_COLOR,
)
country_code_entry.pack(fill="x")

extract_button_outer = tk.Frame(right_col, bg=ACCENT_RED)
extract_button_outer.pack(anchor="w")
extract_button_inner = tk.Frame(extract_button_outer, bg=PRIMARY_BG)
extract_button_inner.pack(padx=2, pady=2)

extract_button = ttk.Button(extract_button_inner, text="Extract Fullnames", style="App.TButton", command=process_ids)
extract_button.pack()

# Treeview section
table_outer, table_wrap = bordered_frame(container, padding=10)
table_outer.pack(fill="both", expand=True)

columns = ("ID", "Fullname", "LocalGroups", "GlobalGroups")
tree = ttk.Treeview(table_wrap, columns=columns, show="headings", height=10, style="App.Treeview")
for col in columns:
    tree.heading(col, text=col)
    # Wider columns for better readability
    if col == "ID":
        tree.column(col, width=120, anchor="w")
    elif col == "Fullname":
        tree.column(col, width=220, anchor="w")
    else:
        tree.column(col, width=260, anchor="w")

# Scrollbars with the same bordered look
scroll_holder = tk.Frame(table_wrap, bg=PRIMARY_BG)
scroll_holder.pack(fill="both", expand=True)

tree_frame = tk.Frame(scroll_holder, bg=PRIMARY_BG)
tree_frame.grid(row=0, column=0, sticky="nsew")

vs_outer = tk.Frame(scroll_holder, bg=ACCENT_RED)
vs_outer.grid(row=0, column=1, sticky="ns", padx=(6, 0))
vs_inner = tk.Frame(vs_outer, bg=PRIMARY_BG)
vs_inner.pack(padx=2, pady=2, fill="y")

hs_outer = tk.Frame(scroll_holder, bg=ACCENT_RED)
hs_outer.grid(row=1, column=0, sticky="ew", pady=(6, 0))
hs_inner = tk.Frame(hs_outer, bg=PRIMARY_BG)
hs_inner.pack(padx=2, pady=2, fill="x")

yscroll = ttk.Scrollbar(vs_inner, orient="vertical", command=tree.yview)
xscroll = ttk.Scrollbar(hs_inner, orient="horizontal", command=tree.xview)
tree.configure(yscrollcommand=yscroll.set, xscrollcommand=xscroll.set)

tree.grid(row=0, column=0, sticky="nsew")
yscroll.pack(fill="y", expand=True)
xscroll.pack(fill="x", expand=True)

scroll_holder.grid_rowconfigure(0, weight=1)
scroll_holder.grid_columnconfigure(0, weight=1)

# Responsive weights
top.grid_columnconfigure(0, weight=3)
top.grid_columnconfigure(1, weight=2)

root.mainloop()