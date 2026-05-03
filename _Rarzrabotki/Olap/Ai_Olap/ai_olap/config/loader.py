"""Load pipeline JSON configs from disk."""
from __future__ import annotations

import json
from pathlib import Path

from ..core.exceptions import ValidationError


def load_pipeline_config(path: str | Path) -> dict:
    """Read a pipelines/*.json file and return the parsed dict."""
    p = Path(path)
    if not p.exists():
        raise ValidationError(f"Pipeline config not found: {p}")
    try:
        return json.loads(p.read_text(encoding="utf-8"))
    except json.JSONDecodeError as exc:
        raise ValidationError(f"Pipeline {p.name} is not valid JSON: {exc}") from exc


def discover_pipelines(pipelines_dir: str | Path = "pipelines") -> list[Path]:
    """Return all *.json files in pipelines_dir, alphabetically sorted."""
    d = Path(pipelines_dir)
    if not d.is_dir():
        return []
    return sorted(d.glob("*.json"))
