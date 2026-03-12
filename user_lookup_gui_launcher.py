import tkinter as tk
from tkinter import ttk, messagebox

class SplashScreen(tk.Toplevel):
    def __init__(self, master, duration=2000):
        super().__init__(master)
        self.title("Loading...")
        self.geometry("300x150")
        label = ttk.Label(self, text="User Lookup Extractor", font=("Arial", 16))
        label.pack(expand=True)
        self.after(duration, self.destroy)

class InputScreen(ttk.Frame):
    def __init__(self, master, on_submit):
        super().__init__(master, padding=20)
        self.on_submit = on_submit
        self.grid(row=0, column=0, sticky="nsew")
        master.title("Extraction Parameters")
        master.geometry("400x350")

        self.entries = {}
        fields = [
            ("RAW_FILE", "Excel file path"),
            ("SHEET_NAME", "Sheet name"),
            ("HOSTNAME_COL", "Hostname column"),
            ("OUTPUT_FILE", "Output file name (.xlsx)"),
            ("NUM_ROWS", "Rows to extract (number or 'all')")
        ]
        for i, (key, label) in enumerate(fields):
            ttk.Label(self, text=label).grid(row=i, column=0, sticky="w", pady=5)
            entry = ttk.Entry(self)
            entry.grid(row=i, column=1, pady=5, sticky="ew")
            self.entries[key] = entry
        self.grid_columnconfigure(1, weight=1)

        submit_btn = ttk.Button(self, text="Continue", command=self.submit)
        submit_btn.grid(row=len(fields), column=0, columnspan=2, pady=15)

    def verify_inputs(self, values):
        return  
    def submit(self):
        values = {k: e.get().strip() for k, e in self.entries.items()}
        if not all(values.values()):
            messagebox.showerror("Input Error", "All fields are required.")
            return
        self.on_submit(values)

class MainApp:
    def __init__(self):
        self.root = tk.Tk()
        self.root.withdraw()  # Hide main window during splash
        self.show_splash()

    def show_splash(self):
        splash = SplashScreen(self.root)
        splash.grab_set()
        splash.wait_window()
        self.root.deiconify()
        self.show_input_screen()

    def show_input_screen(self):
        self.input_frame = InputScreen(self.root, self.on_input_submit)

    def on_input_submit(self, values):
        self.input_frame.destroy()
        # Import UserLookupGUI from user_lookup1.2.py
        from user_lookup12 import UserLookupGUI
        # Set config variables
        import user_lookup12 as ul
        ul.RAW_FILE = values["RAW_FILE"]
        ul.SHEET_NAME = values["SHEET_NAME"]
        ul.HOSTNAME_COL = values["HOSTNAME_COL"]
        ul.OUTPUT_FILE = values["OUTPUT_FILE"]
        ul.NUM_ROWS = values["NUM_ROWS"]
        # Launch UserLookupGUI
        UserLookupGUI(self.root)

    def run(self):
        self.root.mainloop()

if __name__ == "__main__":
    MainApp().run()
