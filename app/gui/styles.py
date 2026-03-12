import tkinter as tk
import tkinter.ttk as ttk

RED_ACCENT = "#C0392B"
RED_LIGHT  = "#E74C3C"
RED_DARK   = "#922B21"
BG_DARK    = "#1E1E1E"
BG_MID     = "#2D2D2D"
BG_LIGHT   = "#3C3C3C"
FG         = "#F0F0F0"
FG_DIM     = "#AAAAAA"


def apply_theme(root: tk.Tk) -> None:
    """Apply the dark + red-accent ttk theme to the given root window."""
    style = ttk.Style(root)
    style.theme_use("clam")

    root.configure(bg=BG_DARK)

    style.configure("TFrame",
                    background=BG_MID)

    style.configure("TLabel",
                    background=BG_MID,
                    foreground=FG,
                    font=("Segoe UI", 9))

    style.configure("Dim.TLabel",
                    background=BG_MID,
                    foreground=FG_DIM,
                    font=("Segoe UI", 9))

    style.configure("TEntry",
                    fieldbackground=BG_DARK,
                    foreground=FG,
                    insertcolor=FG,
                    bordercolor=BG_LIGHT,
                    lightcolor=BG_LIGHT,
                    darkcolor=BG_LIGHT)

    style.configure("TButton",
                    background=BG_LIGHT,
                    foreground=FG,
                    font=("Segoe UI", 9),
                    relief="flat",
                    padding=6)
    style.map("TButton",
              background=[("active", BG_MID), ("disabled", BG_DARK)],
              foreground=[("disabled", FG_DIM)])

    style.configure("Accent.TButton",
                    background=RED_ACCENT,
                    foreground="white",
                    font=("Segoe UI", 9, "bold"),
                    relief="flat",
                    padding=6)
    style.map("Accent.TButton",
              background=[("active", RED_LIGHT), ("disabled", BG_LIGHT)],
              foreground=[("disabled", FG_DIM)])

    style.configure("Treeview.Heading",
                    background=RED_ACCENT,
                    foreground="white",
                    font=("Segoe UI", 9, "bold"),
                    relief="flat")
    style.map("Treeview.Heading",
              background=[("active", RED_DARK)])

    style.configure("Treeview",
                    background=BG_DARK,
                    foreground=FG,
                    fieldbackground=BG_DARK,
                    rowheight=26,
                    font=("Segoe UI", 9))
    style.map("Treeview",
              background=[("selected", RED_LIGHT)],
              foreground=[("selected", "white")])

    style.configure("Vertical.TScrollbar",
                    background=BG_LIGHT,
                    troughcolor=BG_DARK,
                    arrowcolor=FG_DIM,
                    bordercolor=BG_DARK)

    style.configure("red.Horizontal.TProgressbar",
                    troughcolor=BG_DARK,
                    background=RED_ACCENT,
                    thickness=6)

    style.configure("TSeparator",
                    background=RED_ACCENT)


def add_border_accent(widget: tk.Widget, colour: str = RED_ACCENT, thickness: int = 2) -> tk.Frame:
    """
    Wrap a widget in a thin coloured frame to simulate a gradient border accent.
    Returns the outer frame — pack/grid the outer frame instead of the widget.

    Usage:
        inner = ttk.Frame(parent)
        bordered = add_border_accent(inner)
        bordered.pack(fill="both", expand=True, padx=8, pady=8)
    """
    outer = tk.Frame(widget.master, bg=colour, padx=thickness, pady=thickness)
    widget.master = outer
    widget.pack(fill="both", expand=True)
    return outer
