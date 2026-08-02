"""Validate pipeline configs against the JSON schema."""
from __future__ import annotations

import jsonschema

from ..core.exceptions import ValidationError
from .schema import PIPELINE_SCHEMA


def validate_pipeline(config: dict, *, source: str = "<config>") -> None:
    """Raise ValidationError if config does not match PIPELINE_SCHEMA."""
    try:
        jsonschema.validate(config, PIPELINE_SCHEMA)
    except jsonschema.ValidationError as exc:
        path = ".".join(str(p) for p in exc.absolute_path) or "<root>"
        raise ValidationError(f"{source}: {path}: {exc.message}") from exc
