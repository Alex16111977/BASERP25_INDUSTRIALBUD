"""Apply a chain of named transformers to extracted rows.

A pipeline step config:

    "transformer": {
        "steps": ["varbinary_to_uuid", "enum_resolver", "drill_down",
                  "onec_date", "column_mapper"],
        "options": {
            "varbinary_to_uuid": {"columns": ["_IDRRef", "_RecorderRRef", ...]},
            "enum_resolver":     {"column_to_enum": {"Source": "Перечисление.А_ИсточникPL"}},
            "drill_down":        {"source_column": "_RecorderRRef",
                                  "target_column": "Source_Recorder_Url",
                                  "meta_full": "Документ.А_ФинРез_PL"},
            "onec_date":         {"columns": ["_Date_Time"], "mode": "month"},
            "column_mapper":     {"column_map": {"_IDRRef": "PL_Articles_ID", ...}}
        }
    }
"""
from __future__ import annotations

from importlib import import_module
from typing import Callable

from ..core.exceptions import TransformError

REGISTRY: dict[str, str] = {
    "varbinary_to_uuid": "ai_olap.transformers.varbinary_to_uuid",
    "onec_date": "ai_olap.transformers.onec_date",
    "period_offset_fix": "ai_olap.transformers.period_offset_fix",
    "enum_resolver": "ai_olap.transformers.enum_resolver",
    "drill_down": "ai_olap.transformers.drill_down",
    "column_mapper": "ai_olap.transformers.column_mapper",
}


def _resolve(name: str) -> Callable:
    mod_path = REGISTRY.get(name)
    if not mod_path:
        raise TransformError(f"Unknown transformer: {name!r}. Known: {list(REGISTRY)}")
    mod = import_module(mod_path)
    if not hasattr(mod, "transform"):
        raise TransformError(f"{mod_path} has no transform() function")
    return mod.transform


def apply(rows: list[dict], steps: list[str], options: dict | None = None) -> list[dict]:
    options = options or {}
    for step in steps:
        opts = options.get(step) or {}
        rows = _resolve(step)(rows, **opts)
    return rows
