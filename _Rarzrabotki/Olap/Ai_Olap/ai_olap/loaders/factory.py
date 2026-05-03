"""Pick the right Loader for a step config."""
from __future__ import annotations

from ..core.exceptions import ValidationError
from .base import Loader
from .bridge import BridgeLoader
from .dim import DimLoader
from .fact import FactLoader


def make_loader(config: dict) -> Loader:
    mode = config.get("mode")
    if mode == "full_reload":
        # Bridges and Dims share the same TRUNCATE+INSERT semantics.
        if config.get("target_table", "").startswith(("PLArticle_DDS", "Bridge_")):
            return BridgeLoader(config)
        return DimLoader(config)
    if mode == "idempotent_period":
        return FactLoader(config)
    if mode == "append":
        # Generic append (used for ETL_Runs, but we use direct helpers there).
        return DimLoader(config)
    raise ValidationError(f"Unknown loader mode: {mode!r}")
