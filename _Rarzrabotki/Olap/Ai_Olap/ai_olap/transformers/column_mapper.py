"""Rename columns from extractor schema (1С or raw SQL) to OlapBASERP target schema.

The mapping is part of the pipeline step config under transformer.column_map:
    {"_IDRRef": "Organization_ID", "Наименование": "Organization_Name", ...}

Unspecified columns are dropped unless `keep_extra=True`.
"""
from __future__ import annotations


def transform(
    rows: list[dict],
    *,
    column_map: dict[str, str],
    defaults: dict | None = None,
    keep_extra: bool = False,
) -> list[dict]:
    """
    column_map: {src_key_in_extractor_row: dst_key_in_target_table}
    defaults:   {dst_key: value} — written if the dst is not produced by column_map.
    keep_extra: if True, columns not in column_map flow through unchanged.
    """
    if not rows:
        return rows
    defaults = defaults or {}
    out: list[dict] = []
    for row in rows:
        new = {}
        for src, dst in column_map.items():
            if src in row:
                new[dst] = row[src]
        for k, v in defaults.items():
            if new.get(k) is None:
                new[k] = v
        if keep_extra:
            for k, v in row.items():
                if k not in column_map:
                    new[k] = v
        out.append(new)
    return out
