"""Resolve 1С Enum refs to frozen string identifiers.

The Power BI model relies on these strings literally — never rename them.

For each known enum:
1) build a one-shot {hex_uuid: frozen_string} map by reading _IDRRef + _EnumOrder
   from the SQL backend (table from mapping_resolver) and zipping with the
   frozen list defined here;
2) cache the map for the process lifetime;
3) replace UUIDs in input rows.

If a row has a UUID not present in the map, the value becomes None (broken ref)
and a warning is logged.
"""
from __future__ import annotations

import functools

import structlog

from ..core.connections import get_baserp_sql
from ..core.exceptions import TransformError
from ..utils.mapping_resolver import resolve

log = structlog.get_logger().bind(component="enum_resolver")


# Frozen enum value lists — DO NOT REORDER; Power BI DAX references these literals.
FROZEN_ENUMS: dict[str, list[str]] = {
    "Перечисление.А_ИсточникPL": [
        "PL_Excel",
        "ERP_OpEx",
        "ERP_CoGS",
        "ERP_Income",
        "ERP_БезPL_Расх",
        "ERP_БезPL_Доход",
        "Казна_PL",
        "Казна_БезPL",
    ],
    "Перечисление.А_ИсточникDDS": [
        "ERP_Безнал",
        "ERP_Нал",
        "Казна",
    ],
    "Перечисление.А_РазделыCFS": [
        "Operating",
        "Investing",
        "Financing",
        "Internal",
    ],
}


@functools.lru_cache(maxsize=8)
def _load_enum_map(meta_full: str) -> dict[str, str]:
    """Build {hex_uuid_or_bytes_hex: frozen_string} for an enum metadata path."""
    if meta_full not in FROZEN_ENUMS:
        raise TransformError(f"No frozen list defined for {meta_full}")
    table, fields = resolve(meta_full)
    idref = fields["Ссылка"]
    order = fields["Порядок"]
    out: dict[str, str] = {}
    with get_baserp_sql() as c:
        cur = c.cursor()
        cur.execute(f"SELECT {idref}, {order} FROM {table}")
        for row in cur.fetchall():
            uid: bytes = row[0]
            n: int = int(row[1])
            try:
                value = FROZEN_ENUMS[meta_full][n]
            except IndexError as exc:
                raise TransformError(
                    f"{meta_full}: enum order {n} out of frozen list "
                    f"(size {len(FROZEN_ENUMS[meta_full])}). Configuration changed?"
                ) from exc
            out[uid.hex()] = value
    log.info("enum loaded", enum=meta_full, count=len(out))
    return out


def reload_cache() -> None:
    _load_enum_map.cache_clear()


def transform(rows: list[dict], *, column_to_enum: dict[str, str]) -> list[dict]:
    """Replace UUID hex strings (or bytes) with frozen strings.

    column_to_enum: {row_key: 1C_metadata_full_name}, e.g.
        {"Source": "Перечисление.А_ИсточникPL"}
    """
    maps = {col: _load_enum_map(meta) for col, meta in column_to_enum.items()}
    for row in rows:
        for col, m in maps.items():
            if col not in row:
                continue
            val = row[col]
            if val is None:
                continue
            if isinstance(val, (bytes, bytearray, memoryview)):
                val = bytes(val).hex()
            row[col] = m.get(val)
    return rows
