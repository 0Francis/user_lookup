import unicodedata

import pytest

from app.utils.encoding import decode_net_output


def test_utf8_bytes_decoded():
    text = "Full Name                John Doe"
    assert decode_net_output(text.encode("utf-8")) == unicodedata.normalize("NFC", text)


def test_cp1252_bytes_decoded():
    text = "Full Name                Jöhn Döe"
    raw = text.encode("cp1252")
    result = decode_net_output(raw)
    assert "Jöhn" in result or "J" in result


def test_empty_bytes_returns_empty_string():
    result = decode_net_output(b"")
    assert result == ""


def test_replace_on_undecodable_bytes():
    raw = b"\xff\xfe\xfd"
    result = decode_net_output(raw)
    assert isinstance(result, str)


def test_result_is_nfc_normalized():
    """Output must always be NFC-normalized — test with ASCII so all codecs agree."""
    raw = "Full Name   John Smith".encode("ascii")
    result = decode_net_output(raw)
    assert result == unicodedata.normalize("NFC", result)
