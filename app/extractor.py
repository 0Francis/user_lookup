import os
import re
import subprocess
from typing import Callable, Generator

import pandas as pd

from app.config import AppConfig
from app.utils.encoding import decode_net_output


class ExtractionCancelledError(Exception):
    """Raised when the user cancels mid-extraction."""


class UserExtractor:
    """
    Handles all data work: reading the Excel file, extracting AUUIDs,
    and querying Active Directory via `net user /domain`.
    No GUI code lives here.
    """

    AUUID_PATTERN = re.compile(r"(\d{5,})")

    def __init__(self, config: AppConfig) -> None:
        self.config = config
        self._cancelled = False

    def cancel(self) -> None:
        self._cancelled = True

    def _check_cancelled(self) -> None:
        if self._cancelled:
            raise ExtractionCancelledError("Extraction cancelled by user.")

    def load_dataframe(self) -> pd.DataFrame:
        """Read and explode the source Excel file into one hostname per row."""
        df = pd.read_excel(
            self.config.raw_file,
            sheet_name=self.config.sheet_name,
            engine="openpyxl",
        )
        col = self.config.hostname_col
        df = df.assign(**{col: df[col].astype(str).str.split(r"[\n\r]+")})
        df = df.explode(col).reset_index(drop=True)
        df["_auuid"] = df[col].astype(str).str.extract(self.AUUID_PATTERN, expand=False)
        return df

    def extract(
        self,
        progress_cb: Callable[[int, int], None] | None = None,
    ) -> list[tuple[str, str, str]]:
        """
        Run the full extraction.

        Returns a list of (hostname, auuid, full_name) tuples.
        Calls progress_cb(done, total) after each resolved user if provided.
        Raises ExtractionCancelledError if cancel() was called mid-run.
        """
        self._cancelled = False
        df = self.load_dataframe()

        ids = df["_auuid"].dropna().unique()
        limit = self.config.num_rows
        if limit.lower() != "all":
            ids = ids[: int(limit)]

        total = len(ids)
        records: list[tuple[str, str, str]] = []

        for i, auuid in enumerate(ids, 1):
            self._check_cancelled()
            full_name = self.get_user_info(str(auuid))
            hostname_rows = df[df["_auuid"] == auuid][self.config.hostname_col]
            hostname = hostname_rows.iloc[0] if not hostname_rows.empty else ""
            records.append((str(hostname), str(auuid), full_name))
            if progress_cb:
                progress_cb(i, total)

        return records

    def get_user_info(self, auuid: str) -> str:
        """Query AD for the full name of a single AUUID. Returns a string always."""
        try:
            result = subprocess.run(
                ["cmd", "/c", f"net user /domain {auuid}"],
                capture_output=True,
                check=True,
            )
            output = decode_net_output(result.stdout)
            for line in output.splitlines():
                s = line.strip()
                if s.lower().startswith("full name"):
                    parts = re.split(r"\s{2,}", s)
                    if len(parts) > 1:
                        return parts[1].strip()
                    if ":" in s:
                        return s.split(":", 1)[1].strip()
            return "Not found or blank"
        except subprocess.CalledProcessError as e:
            _AD_ERRORS = {
                2: "User not found in domain",
                1: "Not in domain or access denied",
            }
            return _AD_ERRORS.get(e.returncode, f"AD lookup failed (code {e.returncode})")
        except ExtractionCancelledError:
            raise
        except Exception as e:
            return f"Lookup error: {type(e).__name__}"

    def save_results(self, records: list[tuple[str, str, str]]) -> str:
        """
        Write results to an Excel file as a 'user data' sheet.

        - If the configured output file already exists: append/replace the sheet.
        - If it does not exist but the parent directory does: create a new file there.
        - If the parent directory also does not exist: save to the user's Desktop.

        Returns the final path the file was written to.
        """
        df = pd.DataFrame(records, columns=["Hostname", "AUUID", "Full Name"])

        target = self.config.output_file
        if os.path.isfile(target):
            mode = "a"
            extra = {"if_sheet_exists": "replace"}
        elif os.path.isdir(os.path.dirname(os.path.abspath(target))):
            mode = "w"
            extra = {}
        else:
            desktop = os.path.join(os.path.expanduser("~"), "Desktop")
            target = os.path.join(desktop, "user_lookup_results.xlsx")
            mode = "a" if os.path.isfile(target) else "w"
            extra = {"if_sheet_exists": "replace"} if mode == "a" else {}

        with pd.ExcelWriter(target, engine="openpyxl", mode=mode, **extra) as writer:
            df.to_excel(writer, sheet_name="user data", index=False)

        return target
