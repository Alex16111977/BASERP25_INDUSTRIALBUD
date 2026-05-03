"""Convert VARBINARY(16) UUIDs (1С _IDRRef etc.) to char(32) hex strings.

1С stores object refs as 16-byte VARBINARY. OlapBASERP FK columns are char(32),
so the right canonical form is lowercase hex (no dashes).

The transformer scans each row for bytes-typed values and converts them.
A whitelist of column names can be provided to limit conversion.

A NULL UUID (`b"\\x00" * 16` or all-zero) is mapped to None to keep FKs clean.
"""
from __future__ import annotations

ALL_ZERO = b"\x00" * 16


def _convert(value):
    if isinstance(value, (bytes, bytearray, memoryview)):
        b = bytes(value)
        if len(b) == 16:
            if b == ALL_ZERO:
                return None
            return b.hex()
        if len(b) == 1:
            # _Active / _Marked / _Folder — leave as raw boolean below
            return bool(b[0])
    return value


def transform(rows: list[dict], *, columns: list[str] | None = None) -> list[dict]:
    """Convert VARBINARY values to hex/bool for every row.

    columns: if provided, only listed keys are converted; otherwise all keys.
    """
    if not rows:
        return rows
    keys = columns or list(rows[0].keys())
    for row in rows:
        for k in keys:
            if k in row:
                row[k] = _convert(row[k])
    return rows
