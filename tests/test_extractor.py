import re
from unittest.mock import MagicMock, patch

import openpyxl
import pandas as pd
import pytest

from app.config import AppConfig
from app.extractor import ExtractionCancelledError, UserExtractor


@pytest.fixture()
def sample_xlsx(tmp_path):
    """Create a simple xlsx with hostname data for testing."""
    path = tmp_path / "source.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws.append(["Hostname"])
    ws.append(["PC-12345"])
    ws.append(["LAPTOP-67890"])
    ws.append(["DESK-11111\n22222"])
    wb.save(str(path))
    return str(path)


@pytest.fixture()
def config(sample_xlsx, tmp_path):
    return AppConfig(
        raw_file=sample_xlsx,
        sheet_name="Sheet1",
        hostname_col="Hostname",
        output_file=str(tmp_path / "output.xlsx"),
        num_rows="all",
    )


def test_auuid_pattern_matches_5plus_digits():
    pattern = UserExtractor.AUUID_PATTERN
    assert pattern.search("PC-12345").group(1) == "12345"
    assert pattern.search("LAPTOP-1234567").group(1) == "1234567"
    assert pattern.search("NO-DIGITS") is None
    assert pattern.search("AB-1234") is None


def test_load_dataframe_explodes_multiline(config):
    extractor = UserExtractor(config)
    df = extractor.load_dataframe()
    hostnames = df["Hostname"].tolist()
    assert "DESK-11111" in hostnames
    assert "22222" in hostnames


def test_load_dataframe_extracts_auuid_column(config):
    extractor = UserExtractor(config)
    df = extractor.load_dataframe()
    assert "_auuid" in df.columns
    auuids = df["_auuid"].dropna().tolist()
    assert "12345" in auuids
    assert "67890" in auuids


def test_get_user_info_parses_full_name():
    extractor = UserExtractor(AppConfig())
    fake_output = b"Full Name                Jane Smith\nOther Field    value\n"
    with patch("app.extractor.subprocess.run") as mock_run:
        mock_run.return_value = MagicMock(stdout=fake_output)
        with patch("app.extractor.decode_net_output", return_value=fake_output.decode()):
            result = extractor.get_user_info("12345")
    assert result == "Jane Smith"


def test_get_user_info_returns_not_found_when_blank():
    extractor = UserExtractor(AppConfig())
    fake_output = b"Full Name                \nOther Field    value\n"
    with patch("app.extractor.subprocess.run") as mock_run:
        mock_run.return_value = MagicMock(stdout=fake_output)
        with patch("app.extractor.decode_net_output", return_value=fake_output.decode()):
            result = extractor.get_user_info("99999")
    assert result == "Not found or blank"


def test_get_user_info_handles_subprocess_error():
    import subprocess
    extractor = UserExtractor(AppConfig())
    with patch("app.extractor.subprocess.run", side_effect=subprocess.CalledProcessError(1, "net")):
        result = extractor.get_user_info("00000")
    assert "domain" in result.lower() or "denied" in result.lower() or "failed" in result.lower()


def test_cancel_stops_extraction(config):
    extractor = UserExtractor(config)
    calls = []

    def fake_get_user_info(auuid):
        calls.append(auuid)
        extractor.cancel()
        return "Test User"

    extractor.get_user_info = fake_get_user_info
    with pytest.raises(ExtractionCancelledError):
        extractor.extract()
    assert len(calls) >= 1


def test_save_results_writes_xlsx(config, tmp_path):
    wb = openpyxl.Workbook()
    wb.save(config.output_file)
    extractor = UserExtractor(config)
    records = [("PC-12345", "12345", "John Doe"), ("LAPTOP-67890", "67890", "Jane Smith")]
    extractor.save_results(records)
    df = pd.read_excel(config.output_file, sheet_name="user data", engine="openpyxl")
    assert list(df.columns) == ["Hostname", "AUUID", "Full Name"]
    assert len(df) == 2
    assert df.iloc[0]["Full Name"] == "John Doe"
