"""Compose drill-down URLs that 1С Web Client understands.

URL pattern: e1cib/data/<MetaPath>?ref=<32-hex-uuid>
e.g.        e1cib/data/Document.А_ФинРез_PL?ref=80ca00155d2353091e9efe7a6d5b7c81

For typed references in registers (recorders) we get three columns:
  _<fld>_TYPE  - 1 byte, ignored (always 1 = ref)
  _<fld>_RTRef - 4-byte type id (which Document table)
  _<fld>_RRRef - 16-byte UUID

For typed columns the pipeline must supply a fixed metadata path (`recorder_meta`)
because mapping the type id back to a metadata name requires a separate DBNames
lookup; we keep it simple — facts are bound to their recorder type by design.
"""
from __future__ import annotations


def build_url(meta_full: str, uuid_hex: str | bytes | None) -> str | None:
    if uuid_hex is None:
        return None
    if isinstance(uuid_hex, (bytes, bytearray, memoryview)):
        uuid_hex = bytes(uuid_hex).hex()
    if uuid_hex == "00" * 16:
        return None
    return f"e1cib/data/{meta_full}?ref={uuid_hex}"


def transform(
    rows: list[dict],
    *,
    source_column: str,
    target_column: str,
    meta_full: str,
) -> list[dict]:
    """For every row, compute a drill-down URL and store it under target_column."""
    for row in rows:
        row[target_column] = build_url(meta_full, row.get(source_column))
    return rows
