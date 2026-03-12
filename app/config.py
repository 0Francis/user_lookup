import json
import os
import sys
from dataclasses import asdict, dataclass, field
from typing import Optional


def resource_path(relative: str) -> str:
    """Resolve a path that works both in development and when frozen by PyInstaller."""
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, relative)


def _config_path() -> str:
    """
    Path to settings.json — written next to the executable (or project root
    in dev) so it persists between runs without being bundled inside the exe.
    """
    if getattr(sys, "frozen", False):
        base = os.path.dirname(sys.executable)
    else:
        base = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    config_dir = os.path.join(base, "config")
    os.makedirs(config_dir, exist_ok=True)
    return os.path.join(config_dir, "settings.json")


@dataclass
class AppConfig:
    raw_file: str = ""
    sheet_name: str = ""
    hostname_col: str = ""
    output_file: str = ""
    num_rows: str = "all"

    def validate(self) -> list[str]:
        """Return a list of validation error messages. Empty list means valid."""
        errors: list[str] = []
        if not self.raw_file:
            errors.append("raw_file is required.")
        elif not os.path.isfile(self.raw_file):
            errors.append(f"raw_file not found: {self.raw_file}")
        if not self.sheet_name:
            errors.append("sheet_name is required.")
        if not self.hostname_col:
            errors.append("hostname_col is required.")
        if not self.output_file:
            errors.append("output_file is required.")
        elif not self.output_file.endswith(".xlsx"):
            errors.append("output_file must end with .xlsx")
        if self.num_rows.lower() != "all":
            if not self.num_rows.isdigit() or int(self.num_rows) < 1:
                errors.append("num_rows must be a positive integer or 'all'.")
        return errors

    def save(self) -> None:
        path = _config_path()
        with open(path, "w", encoding="utf-8") as f:
            json.dump(asdict(self), f, indent=2)

    @classmethod
    def load(cls) -> "AppConfig":
        path = _config_path()
        if not os.path.isfile(path):
            return cls()
        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
            return cls(**{k: v for k, v in data.items() if k in cls.__dataclass_fields__})
        except Exception:
            return cls()
