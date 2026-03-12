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
            return f"AD error: {e}"
        except ExtractionCancelledError:
            raise
        except Exception as e:
            return f"Unexpected error: {e}"

    def save_results(self, records: list[tuple[str, str, str]]) -> None:
        """Append results to the configured output Excel file as a new sheet."""
        df = pd.DataFrame(records, columns=["Hostname", "AUUID", "Full Name"])
        with pd.ExcelWriter(
            self.config.output_file, engine="openpyxl", mode="a", if_sheet_exists="replace"
        ) as writer:
            df.to_excel(writer, sheet_name="user data", index=False)
