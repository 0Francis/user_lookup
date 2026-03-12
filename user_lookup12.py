import pandas as pd
import subprocess
import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox
import re
import threading
import ctypes
import locale
import unicodedata

# CONFIG
RAW_FILE = r""
SHEET_NAME = ""
HOSTNAME_COL = ""
OUTPUT_FILE = ""
NUM_ROWS = ""
OUTPUT_SHEET = "user data"

class UserLookupGUI:
    def __init__(self, master):
        self.master = master
        master.title("User Lookup Extract Viewer")
        self.frame = ttk.Frame(master, padding=10)
        self.frame.pack(fill=tk.BOTH, expand=True)
        
        self.start_btn = ttk.Button(self.frame, text="Start Extraction", command=self.start_process)
        self.start_btn.grid(row=0, column=0, padx=5, pady=5)
        
        self.cancel_btn = ttk.Button(self.frame, text="Cancel", command=self.cancel_process, state=tk.DISABLED)
        self.cancel_btn.grid(row=0, column=1, padx=5, pady=5)
        
        # Add Treeview with Scrollbar
        self.tree_frame = ttk.Frame(self.frame)
        self.tree_frame.grid(row=1, column=0, columnspan=2, sticky="nsew")
        self.tree = ttk.Treeview(self.tree_frame, columns=("Hostname", "AUUID", "Full Name"), show="headings", height=20)
        self.tree.heading("Hostname", text="Hostname")
        self.tree.heading("AUUID", text="AUUID")
        self.tree.heading("Full Name", text="Full Name")
        self.tree.grid(row=0, column=0, sticky="nsew")
        
        self.vsb = ttk.Scrollbar(self.tree_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=self.vsb.set)
        self.vsb.grid(row=0, column=1, sticky="ns")
        
        self.tree_frame.rowconfigure(0, weight=1)
        self.tree_frame.columnconfigure(0, weight=1)
        
        self.frame.rowconfigure(1, weight=1)
        self.frame.columnconfigure(0, weight=1)
        self.frame.columnconfigure(1, weight=1)
        
        self.load_btn = ttk.Button(self.frame, text="Load Data", command=self.load_data)
        self.load_btn.grid(row=2, column=0, columnspan=2, pady=5)
        
        self.cancel_requested = False
    
    def _decode_net_output(self, raw: bytes) -> str:
        import ctypes, locale, unicodedata
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
    
    def start_process(self):
        self.start_btn.config(state=tk.DISABLED)
        self.cancel_btn.config(state=tk.NORMAL)
        self.cancel_requested = False
        threading.Thread(target=self.run_extraction, daemon=True).start()
    
    def cancel_process(self):
        self.cancel_requested = True
        self.cancel_btn.config(state=tk.DISABLED)
    
    def load_data(self):
        try:
            #Writing/appending data the output to an excel file
            df = pd.DataFrame([self.tree.item(row)["values"] for row in self.tree.get_children()], columns=["Hostname", "AUUID", "Full Name"])
            with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl", mode="a") as writer:
                df.to_excel(writer, sheet_name=OUTPUT_SHEET, index=False)
                
        except Exception as e:
            messagebox.showerror("Error", str(e))
    
    def run_extraction(self):
        try:
            df = pd.read_excel(RAW_FILE, sheet_name=SHEET_NAME, engine="openpyxl")
            df = df.assign(**{HOSTNAME_COL: df[HOSTNAME_COL].astype(str).str.split(r'[\n\r]+')})
            df = df.explode(HOSTNAME_COL).reset_index(drop=True)
            df["AUUID_digits"] = df[HOSTNAME_COL].astype(str).str.extract(r"(\d{5,})", expand=False)
            ids = df["AUUID_digits"].dropna().unique()
            records = []
            for i, auuid in enumerate(ids[:20], 1):
                if self.cancel_requested:
                    break
                info = self.get_user_info(str(auuid))
                hostname_rows = df[df["AUUID_digits"] == auuid][HOSTNAME_COL]
                hostname = hostname_rows.iloc[0] if not hostname_rows.empty else ""
                records.append((hostname, auuid, info))
            self.display_records(records)
        except Exception as e:
            messagebox.showerror("Error", str(e))
        finally:
            self.start_btn.config(state=tk.NORMAL)
            self.cancel_btn.config(state=tk.DISABLED)
    
    def get_user_info(self, auuid):
        try:
            result = subprocess.run(
                ["cmd", "/c", f"net user /domain {str(auuid)}"],
                capture_output=True,
                check=True,
            )
            output = self._decode_net_output(result.stdout)
            fullname = ""
            for line in output.splitlines():
                s = line.strip()
                if s.lower().startswith("full name"):
                    # Try both double-space and colon split
                    parts = re.split(r"\s{2,}", s)
                    if len(parts) > 1:
                        fullname = parts[1].strip()
                    elif ':' in s:
                        fullname = s.split(':', 1)[1].strip()
                    break
            if not fullname:
                return "Not found or blank"
            return fullname
        except subprocess.CalledProcessError as e:
            return f"Error: {e}"
        except Exception as ex:
            return f"Unexpected error: {ex}"
    
    def display_records(self, records):
        for row in self.tree.get_children():
            self.tree.delete(row)
        for rec in records:
            self.tree.insert("", "end", values=rec)
    
def launch_gui():
    root = tk.Tk()
    app = UserLookupGUI(root)
    root.mainloop()
launch_gui()