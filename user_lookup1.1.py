import pandas as pd
import subprocess
import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox
import os
import re

# CONFIG
RAW_FILE = r"c:\Users\23225632\Downloads\Kenya Offrole & CWK Dump_27 FEB.xlsx"
SHEET_NAME = "Sheet1"
HOSTNAME_COL = "Hostname"
OUTPUT_FILE = "output.xlsx"

df = pd.read_excel(RAW_FILE, sheet_name=SHEET_NAME, engine = "openpyxl")
df["AUUID_digits"] = df["Hostname"].astype(str).str.extract(r"(\d{5,})").astype(float).astype("Int64")

def _decode_net_output(raw: bytes) -> str:
    try:
        cp = ctypes.windll.kernel32.GetConsoleOutputCP()
        if cp:
            enc = f"cp{cp}"
            text = raw.decode(enc, errors="strict")
            return unicodedata.normalize("NFC", text)
    except Exception:
        pass
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
    try:
        text = raw.decode("mbcs", errors="replace")
    except Exception:
        text = raw.decode(errors="replace")
    return unicodedata.normalize("NFC", text)

def get_user_info(user_id):
    try:
        result = subprocess.run([
            "cmd", "/c", f"net user /domain {user_id}"
        ], capture_output=True, check=True)
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
                for next_line in lines[i + 1:]:
                    ns = next_line.strip()
                    if (
                        ns == "" or
                        ns.startswith("Global Group memberships") or
                        ns.startswith("The command completed successfully")
                    ):
                        break
                    group_names = [g for g in ns.split() if g.startswith("*")]
                    local_groups.extend(group_names)
            if s.startswith("Global Group memberships"):
                parts = re.split(r"\s{2,}", s)
                if len(parts) > 1:
                    global_groups.extend(parts[1:])
                for next_line in lines[i + 1:]:
                    ns = next_line.strip()
                    if ns == "" or ns.startswith("The command completed successfully"):
                        break
                    group_names = [g for g in ns.split() if g.startswith("*")]
                    global_groups.extend(group_names)
        local_groups = [g.replace("*", "").strip() for g in local_groups if g.strip()]
        global_groups = [g.replace("*", "").strip() for g in global_groups if g.strip()]
        return fullname, local_groups, global_groups
    except Exception:
        return "", [], []

# For each ID in the AUUID_digits column, run net user and collect info
results = []
for user_id in df["AUUID_digits"].dropna().astype(str).unique():
    fullname, local_groups, global_groups = get_user_info(user_id)
    results.append({
        "ID": user_id,
        "Fullname": fullname,
        "LocalGroups": ", ".join(local_groups),
        "GlobalGroups": ", ".join(global_groups)
    })

import pandas as pd
results_df = pd.DataFrame(results)
results_df