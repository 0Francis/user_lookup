# Test Suite - User Lookup Extractor

## How to Run

```powershell
# From the project root, with venv activated
venv\Scripts\activate
pytest tests\ -v
```

Expected result: **21 passed** in roughly 3-10 seconds.

To run a single file:
```powershell
pytest tests\test_config.py -v
pytest tests\test_encoding.py -v
pytest tests\test_extractor.py -v
```

---

## Test Files

### `test_config.py` - 8 tests

Covers `AppConfig` in `app/config.py`.

| Test | What it checks |
|---|---|
| `test_default_config_is_empty` | Fresh `AppConfig()` has blank fields and `num_rows="all"` |
| `test_validate_returns_errors_on_blank` | `validate()` returns at least one error when all fields are empty |
| `test_validate_passes_with_valid_data` | No errors when all fields are correctly filled with real files |
| `test_validate_rejects_missing_raw_file` | Error returned when `raw_file` path does not exist on disk |
| `test_validate_rejects_bad_num_rows` | Error returned when `num_rows` is non-numeric and not `"all"` |
| `test_validate_rejects_non_xlsx_output` | Error returned when `output_file` does not end with `.xlsx` |
| `test_save_and_load_roundtrip` | `save()` then `load()` returns identical field values |
| `test_load_returns_default_when_file_missing` | `load()` returns a blank default config if `settings.json` does not exist |

---

### `test_encoding.py` - 5 tests

Covers `decode_net_output` in `app/utils/encoding.py`.

| Test | What it checks |
|---|---|
| `test_utf8_bytes_decoded` | UTF-8 encoded bytes are decoded correctly |
| `test_cp1252_bytes_decoded` | Windows cp1252 encoded bytes (accented characters) are decoded without error |
| `test_empty_bytes_returns_empty_string` | `b""` returns `""` without raising |
| `test_replace_on_undecodable_bytes` | Garbage bytes that cannot be decoded still return a `str` (no crash) |
| `test_result_is_nfc_normalized` | Output is always NFC-normalized Unicode |

---

### `test_extractor.py` - 8 tests

Covers `UserExtractor` in `app/extractor.py`.

| Test | What it checks |
|---|---|
| `test_auuid_pattern_matches_5plus_digits` | Regex extracts 5+ digit sequences and ignores 4-digit ones |
| `test_load_dataframe_explodes_multiline` | Cells with newline-separated hostnames are split into individual rows |
| `test_load_dataframe_extracts_auuid_column` | `_auuid` column is populated with digits extracted from hostnames |
| `test_get_user_info_parses_full_name` | `"Full Name   Jane Smith"` line is parsed into `"Jane Smith"` |
| `test_get_user_info_returns_not_found_when_blank` | Blank Full Name line returns `"Not found or blank"` |
| `test_get_user_info_handles_subprocess_error` | `CalledProcessError` from `net user` returns an `"AD error: ..."` string, does not raise |
| `test_cancel_stops_extraction` | Calling `cancel()` mid-extraction causes `ExtractionCancelledError` to be raised |
| `test_save_results_writes_xlsx` | `save_results()` writes a `"user data"` sheet with correct columns and rows |

---

## What to Expect

```
tests/test_config.py::test_default_config_is_empty           PASSED
tests/test_config.py::test_validate_returns_errors_on_blank  PASSED
tests/test_config.py::test_validate_passes_with_valid_data   PASSED
tests/test_config.py::test_validate_rejects_missing_raw_file PASSED
tests/test_config.py::test_validate_rejects_bad_num_rows     PASSED
tests/test_config.py::test_validate_rejects_non_xlsx_output  PASSED
tests/test_config.py::test_save_and_load_roundtrip           PASSED
tests/test_config.py::test_load_returns_default_when_file_missing PASSED
tests/test_encoding.py::test_utf8_bytes_decoded              PASSED
tests/test_encoding.py::test_cp1252_bytes_decoded            PASSED
tests/test_encoding.py::test_empty_bytes_returns_empty_string PASSED
tests/test_encoding.py::test_replace_on_undecodable_bytes    PASSED
tests/test_encoding.py::test_result_is_nfc_normalized        PASSED
tests/test_extractor.py::test_auuid_pattern_matches_5plus_digits PASSED
tests/test_extractor.py::test_load_dataframe_explodes_multiline  PASSED
tests/test_extractor.py::test_load_dataframe_extracts_auuid_column PASSED
tests/test_extractor.py::test_get_user_info_parses_full_name     PASSED
tests/test_extractor.py::test_get_user_info_returns_not_found_when_blank PASSED
tests/test_extractor.py::test_get_user_info_handles_subprocess_error     PASSED
tests/test_extractor.py::test_cancel_stops_extraction        PASSED
tests/test_extractor.py::test_save_results_writes_xlsx       PASSED

21 passed
```

No network or domain access is required. All AD calls are mocked using `unittest.mock.patch`.
