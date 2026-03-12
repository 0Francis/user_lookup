import json
import os
import tempfile

import pytest

from app.config import AppConfig


@pytest.fixture()
def valid_xlsx(tmp_path):
    """Create a minimal real xlsx file for path-existence checks."""
    import openpyxl
    path = tmp_path / "data.xlsx"
    wb = openpyxl.Workbook()
    wb.save(str(path))
    return str(path)


def test_default_config_is_empty():
    cfg = AppConfig()
    assert cfg.raw_file == ""
    assert cfg.num_rows == "all"


def test_validate_returns_errors_on_blank():
    cfg = AppConfig()
    errors = cfg.validate()
    assert len(errors) > 0


def test_validate_passes_with_valid_data(valid_xlsx, tmp_path):
    out = str(tmp_path / "output.xlsx")
    cfg = AppConfig(
        raw_file=valid_xlsx,
        sheet_name="Sheet",
        hostname_col="Host",
        output_file=out,
        num_rows="all",
    )
    errors = cfg.validate()
    assert errors == []


def test_validate_rejects_missing_raw_file(tmp_path):
    out = str(tmp_path / "output.xlsx")
    cfg = AppConfig(
        raw_file="/nonexistent/path.xlsx",
        sheet_name="Sheet",
        hostname_col="Host",
        output_file=out,
        num_rows="all",
    )
    errors = cfg.validate()
    assert any("raw_file" in e for e in errors)


def test_validate_rejects_bad_num_rows(valid_xlsx, tmp_path):
    out = str(tmp_path / "output.xlsx")
    cfg = AppConfig(
        raw_file=valid_xlsx,
        sheet_name="Sheet",
        hostname_col="Host",
        output_file=out,
        num_rows="abc",
    )
    errors = cfg.validate()
    assert any("num_rows" in e for e in errors)


def test_validate_rejects_non_xlsx_output(valid_xlsx, tmp_path):
    cfg = AppConfig(
        raw_file=valid_xlsx,
        sheet_name="Sheet",
        hostname_col="Host",
        output_file=str(tmp_path / "output.csv"),
        num_rows="all",
    )
    errors = cfg.validate()
    assert any(".xlsx" in e for e in errors)


def test_save_and_load_roundtrip(tmp_path, monkeypatch):
    config_dir = tmp_path / "config"
    config_dir.mkdir()
    monkeypatch.setattr(
        "app.config._config_path",
        lambda: str(config_dir / "settings.json"),
    )
    cfg = AppConfig(
        raw_file="C:/some/file.xlsx",
        sheet_name="Sheet1",
        hostname_col="Hostname",
        output_file="C:/some/output.xlsx",
        num_rows="50",
    )
    cfg.save()
    loaded = AppConfig.load()
    assert loaded.raw_file == cfg.raw_file
    assert loaded.sheet_name == cfg.sheet_name
    assert loaded.num_rows == cfg.num_rows


def test_load_returns_default_when_file_missing(tmp_path, monkeypatch):
    monkeypatch.setattr(
        "app.config._config_path",
        lambda: str(tmp_path / "config" / "settings.json"),
    )
    cfg = AppConfig.load()
    assert cfg.raw_file == ""
