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

Dynamic enums (DYNAMIC_ENUMS): for enums that are too large to maintain frozen
(>100 values, e.g. ХозяйственныеОперации has ~600 values), build the map by reading
_IDRRef + _Description directly from the SQL backend. No frozen list required.
"""
from __future__ import annotations

import functools

import structlog

from ..core.connections import get_baserp_sql
from ..core.exceptions import TransformError
from ..utils.mapping_resolver import resolve

log = structlog.get_logger().bind(component="enum_resolver")


# Frozen enum value lists — DO NOT REORDER; Power BI DAX references these literals.
# Order MUST match _EnumOrder values in 1C metadata.
FROZEN_ENUMS: dict[str, list[str]] = {
    "Перечисление.А_ИсточникPL": [
        "PL_Excel",
        "PL_ЕРП",
    ],
    # Stage v3 (2026-05-08): rebuild А_ИсточникDDS to 4 values matching new register
    "Перечисление.А_ИсточникDDS": [
        "ЕРП",          # order 0
        "Казна",        # order 1
        "План",         # order 2
        "ПланОбъекта",  # order 3
    ],
    "Перечисление.А_РазделыCFS": [
        "Operating",
        "Investing",
        "Financing",
        "Internal",
    ],
    "Перечисление.ТипыДвиженияДенежныхСредств": [
        "Поступление",  # order 0
        "Списание",     # order 1
    ],
    "Перечисление.ТипыДенежныхСредств": [
        "Наличные",                          # order 0
        "Безналичные",                       # order 1
        "ДенежныеСредстваУЭквайера",         # order 2
        "ДенежныеСредстваУПодотчетногоЛица", # order 3
        "Депозиты",                          # order 4
        "ДенежныеСредстваВПути",             # order 5
        "ДенежныеДокументы",                 # order 6
    ],
    # Balance Stage (2026-05-16): Fact_Balance Source. Order MUST match
    # Enums/А_ИсточникБаланса.xml EnumValue sequence (_EnumOrder 0..6).
    "Перечисление.А_ИсточникБаланса": [
        "ПрочиеАктивыПассивы",         # order 0
        "РасчетыСКлиентами",           # order 1
        "РасчетыСПоставщиками",        # order 2
        "СебестоимостьТоваров",        # order 3
        "ДенежныеСредстваБезналичные", # order 4
        "ДенежныеСредстваНаличные",    # order 5
        "ПрочиеРасходы",               # order 6
    ],
    # Dim_PAP_Articles.AktivPassiv (канон OD-9). ASCII identifiers — referenced
    # literally by PL.pbix DAX ([AktivPassiv]="Aktiv"|"Passiv"|"AktivPassiv")
    # and stored in varchar(15). Order MUST match
    # Enums/ВидыСтатейУправленческогоБаланса.xml (_EnumOrder 0..2).
    "Перечисление.ВидыСтатейУправленческогоБаланса": [
        "Aktiv",        # order 0 — Актив
        "Passiv",       # order 1 — Пассив
        "AktivPassiv",  # order 2 — Актив/Пассив
    ],
}


# Dynamic enums — too large for frozen list. Resolver reads _Description directly.
# Use 1C metadata name (synonym) as the string identifier.
DYNAMIC_ENUMS: set[str] = {
    "Перечисление.ХозяйственныеОперации",
}


@functools.lru_cache(maxsize=16)
def _load_enum_map(meta_full: str) -> dict[str, str]:
    """Build {hex_uuid: frozen_string} for an enum metadata path.

    For frozen enums — zip _EnumOrder with FROZEN_ENUMS[meta_full].
    For dynamic enums — read _Description directly (no frozen list).
    """
    if meta_full in FROZEN_ENUMS:
        return _load_frozen_enum_map(meta_full)
    elif meta_full in DYNAMIC_ENUMS:
        return _load_dynamic_enum_map(meta_full)
    else:
        raise TransformError(f"No enum mapping defined for {meta_full}")


def _load_frozen_enum_map(meta_full: str) -> dict[str, str]:
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
    log.info("frozen enum loaded", enum=meta_full, count=len(out))
    return out


def _load_dynamic_enum_map(meta_full: str) -> dict[str, str]:
    """Read all UUID -> Description (Synonym/Name) for large enums.

    1C platform stores enum value name in `_Description` column of the _Enum table.
    """
    table, fields = resolve(meta_full)
    idref = fields["Ссылка"]
    out: dict[str, str] = {}
    with get_baserp_sql() as c:
        cur = c.cursor()
        # _Description holds the enum value Synonym (display name)
        cur.execute(f"SELECT {idref}, _Description FROM {table}")
        for row in cur.fetchall():
            uid: bytes = row[0]
            descr: str = (row[1] or "").strip()
            out[uid.hex()] = descr
    log.info("dynamic enum loaded", enum=meta_full, count=len(out))
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
