import ctypes
import locale
import unicodedata


def decode_net_output(raw: bytes) -> str:
    """
    Decode raw bytes from a Windows `net user` subprocess call.

    Strategy:
    1. Try the active console output codepage (GetConsoleOutputCP).
    2. Fall through a prioritised list of common Windows encodings.
    3. Last resort: decode with 'replace' so we never raise.
    """
    try:
        cp = ctypes.windll.kernel32.GetConsoleOutputCP()
        if cp:
            text = raw.decode(f"cp{cp}", errors="strict")
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
    tried: set[str] = set()
    for enc in fallbacks:
        if not enc:
            continue
        enc_lower = enc.lower()
        if enc_lower in tried:
            continue
        tried.add(enc_lower)
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
