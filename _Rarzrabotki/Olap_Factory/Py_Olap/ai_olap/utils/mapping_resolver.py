"""Resolve 1С metadata names to MSSQL backend table/column names.

Reads mapping/baserp_storage.json (produced by mapping/refresh_mapping.py).
"""
from __future__ import annotations

import functools
import json
import os
from pathlib import Path

from ..core.exceptions import MappingError


@functools.lru_cache(maxsize=1)
def _load_mapping(path: str) -> dict:
    p = Path(path)
    if not p.exists():
        raise MappingError(
            f"mapping/baserp_storage.json not found at {p}. "
            f"Run `python mapping/refresh_mapping.py` first."
        )
    try:
        data = json.loads(p.read_text(encoding="utf-8"))
    except json.JSONDecodeError as exc:
        raise MappingError(f"baserp_storage.json malformed: {exc}") from exc
    if "objects" not in data:
        raise MappingError("baserp_storage.json missing top-level 'objects'")
    return data


def _default_path() -> str:
    return os.environ.get("MAPPING_PATH") or "mapping/baserp_storage.json"


def resolve(meta_full_name: str, *, mapping_path: str | None = None) -> tuple[str, dict[str, str]]:
    """Return (primary_table, {1C_field_name: SQL_column_name}) for a 1С object.

    Example:
        resolve("Справочник.Организации")
        -> ("_Reference329", {"Ссылка": "_IDRRef", "Код": "_Code", "Наименование": "_Description", ...})
    """
    data = _load_mapping(mapping_path or _default_path())
    obj = data["objects"].get(meta_full_name)
    if obj is None:
        raise MappingError(
            f"Object {meta_full_name!r} is not in mapping. "
            f"Add it to WHITELIST in mapping/refresh_mapping.py and re-run."
        )
    table = obj.get("primary_table")
    fields = obj.get("fields", {})
    if not table:
        raise MappingError(f"{meta_full_name}: primary_table missing in mapping")
    return table, dict(fields)


def resolve_table_only(meta_full_name: str, *, mapping_path: str | None = None) -> str:
    """Return only the primary SQL table name for a 1С object."""
    return resolve(meta_full_name, mapping_path=mapping_path)[0]


def list_objects(*, mapping_path: str | None = None) -> list[str]:
    data = _load_mapping(mapping_path or _default_path())
    return sorted(data["objects"].keys())


def reload_cache() -> None:
    """Drop the lru_cache (used after mapping/refresh_mapping.py)."""
    _load_mapping.cache_clear()
