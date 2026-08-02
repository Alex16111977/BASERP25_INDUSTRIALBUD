"""Pick the right extractor implementation for a step config."""
from __future__ import annotations

from ..core.exceptions import ValidationError
from .base import Extractor
from .com import ComExtractor
from .sql_backend import SqlBackendExtractor


def make_extractor(config: dict) -> Extractor:
    t = config.get("type")
    if t in ("sql_backend", "raw_sql"):
        return SqlBackendExtractor(config)
    if t == "com":
        return ComExtractor(config)
    raise ValidationError(f"Unknown extractor type: {t!r}")
