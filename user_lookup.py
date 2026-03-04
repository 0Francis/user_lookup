import pandas as pd
import subprocess
import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox
import os
import re

# CONFIG
RAW_FILE = r"c:\Users\23225632\Downloads\Kenya Offrole & CWK Dump_27 FEB.xlsx"
SHEET_NAME = "Sheet5"
HOSTNAME_COL = "Hostname"
OUTPUT_FILE = "output.xlsx"

def extract_valid_ids(df, col):
    df = df[df[col].notnull() & (df[col].astype(str).str.strip() != "")]
    df[col] = df[col].astype(str).str.extract(r"(\d+)", expand=False)
    return df[df[col].notnull() & df[col].astype(str).str.match(r"^\d{6,}$")][col].astype(str).unique()

def get_user_info(user_id):
    try:
        result = subprocess.run(
            ["net", "user", "/domain", user_id],
            capture_output=True,
            text=True,
            check=True,
            encoding="utf-8",
            errors="replace"
        )
        output = result.stdout
        fullname = ""
        local_groups = []
        global_groups = []
        lines = output.splitlines()
        for i, line in enumerate(lines):
            if line.strip().startswith("Full Name"):
                parts = re.split(r"\s{2,}", line.strip())
                if len(parts) > 1:
                    fullname = parts[-1].strip()
            if line.strip().startswith("Local Group Memberships"):
                parts = re.split(r"\s{2,}", line.strip())
                if len(parts) > 1:
                    local_groups.extend(parts[1:])
                for next_line in lines[i+1:]:
                    if next_line.strip() == "" or next_line.strip().startswith("Global Group memberships") or next_line.strip().startswith("The command completed successfully"):
                        break
                    group_names = [g for g in next_line.strip().split() if g.startswith("*")]
                    local_groups.extend(group_names)
            if line.strip().startswith("Global Group memberships"):
                parts = re.split(r"\s{2,}", line.strip())
                if len(parts) > 1:
                    global_groups.extend(parts[1:])
                for next_line in lines[i+1:]:
                    if next_line.strip() == "" or next_line.strip().startswith("The command completed successfully"):
                        break
                    group_names = [g for g in next_line.strip().split() if g.startswith("*")]
                    global_groups.extend(group_names)
        local_groups = [g.replace("*", "").strip() for g in local_groups]
        global_groups = [g.replace("*", "").strip() for g in global_groups]
        return fullname, local_groups, global_groups
    except Exception:
        return "", [], []

def process_ids():
    ids = text_area.get("1.0", tk.END).strip().splitlines()
    # Accept multiple country codes separated by comma or space
    raw_codes = country_code_var.get().strip().upper()
    selected_codes = [code.strip() for code in re.split(r",|\s", raw_codes) if code.strip()]
    results = []
    tree.delete(*tree.get_children())
    for i, user_id in enumerate(ids):
        fullname, local_groups, global_groups = get_user_info(user_id)
        if not selected_codes:
            # No country code input, show all
            results.append({
                "ID": user_id,
                "Fullname": fullname,
                "LocalGroups": ", ".join(local_groups),
                "GlobalGroups": ", ".join(global_groups)
            })
            tree.insert("", "end", values=(
                user_id,
                fullname,
                ", ".join(local_groups),
                ", ".join(global_groups)
            ))
        else:
            # Check if any selected code is in any group
            def has_code(groups):
                return any(any(code in g for code in selected_codes) for g in groups)
            if has_code(local_groups) or has_code(global_groups):
                filtered_local = [g for g in local_groups if any(code in g for code in selected_codes)]
                filtered_global = [g for g in global_groups if any(code in g for code in selected_codes)]
                results.append({
                    "ID": user_id,
                    "Fullname": fullname,
                    "LocalGroups": ", ".join(filtered_local),
                    "GlobalGroups": ", ".join(filtered_global)
                })
                tree.insert("", "end", values=(
                    user_id,
                    fullname,
                    ", ".join(filtered_local),
                    ", ".join(filtered_global)
                ))
        root.update()
    out_df = pd.DataFrame(results)
    out_df.to_excel(OUTPUT_FILE, index=False, engine="openpyxl")
    messagebox.showinfo("Done", f"Results saved to {OUTPUT_FILE}")

root = tk.Tk()
root.title("Domain Fullname Extractor")

# Paste IDs label and text area
label_ids = tk.Label(root, text="Paste IDs (one per line):")
label_ids.pack()
text_area = scrolledtext.ScrolledText(root, width=40, height=10, borderwidth=1, relief="solid")
text_area.pack(pady=(0, 8))

# Country code label and entry (thinner border, smaller width)
frame_country = tk.Frame(root)
frame_country.pack(pady=(0, 8))
label_country = tk.Label(frame_country, text="Country Code(s):")
label_country.pack(side=tk.LEFT)
country_code_var = tk.StringVar()
country_code_entry = tk.Entry(frame_country, textvariable=country_code_var, width=10, borderwidth=1, relief="solid")
country_code_entry.pack(side=tk.LEFT, padx=(4, 0))
label_hint = tk.Label(frame_country, text="(e.g. KE, RW, TZ)", font=("Arial", 8))
label_hint.pack(side=tk.LEFT, padx=(6, 0))

# Button with proper space
button_frame = tk.Frame(root)
button_frame.pack(pady=(0, 10))
extract_btn = tk.Button(button_frame, text="Extract Fullnames", command=process_ids, width=18)
extract_btn.pack()

columns = ("ID", "Fullname", "LocalGroups", "GlobalGroups")
tree_frame = tk.Frame(root)
tree_frame.pack(pady=(0, 8), fill=tk.BOTH, expand=True)
tree_scroll_y = tk.Scrollbar(tree_frame, orient=tk.VERTICAL)
tree_scroll_x = tk.Scrollbar(tree_frame, orient=tk.HORIZONTAL)
tree = ttk.Treeview(tree_frame, columns=columns, show="headings", height=20, yscrollcommand=tree_scroll_y.set, xscrollcommand=tree_scroll_x.set)
tree_scroll_y.config(command=tree.yview)
tree_scroll_x.config(command=tree.xview)
tree_scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
tree_scroll_x.pack(side=tk.BOTTOM, fill=tk.X)
tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
for col in columns:
    tree.heading(col, text=col)
    tree.column(col, width=150)
tree.pack(pady=(0, 8))

root.mainloop()