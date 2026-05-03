"""JSON Schema for pipeline configs (pipelines/*.json)."""

PIPELINE_SCHEMA: dict = {
    "$schema": "https://json-schema.org/draft/2020-12/schema",
    "type": "object",
    "required": ["pipeline_id", "steps"],
    "additionalProperties": False,
    "properties": {
        "pipeline_id": {"type": "string", "pattern": r"^[a-z][a-z0-9_]*$"},
        "description": {"type": "string"},
        "schedule": {"type": ["string", "null"], "description": "Cron expression"},
        "steps": {
            "type": "array",
            "minItems": 1,
            "items": {
                "type": "object",
                "required": ["step_id", "extractor", "loader"],
                "additionalProperties": False,
                "properties": {
                    "step_id": {"type": "string"},
                    "extractor": {
                        "type": "object",
                        "required": ["type"],
                        "properties": {
                            "type": {"enum": ["sql_backend", "com", "raw_sql"]},
                            "object": {"type": "string"},
                            "fields": {"type": "array", "items": {"type": "string"}},
                            "filters": {"type": "object"},
                            "raw_sql": {"type": "string"},
                            "params": {"type": "object"},
                            "com_query": {"type": "string"},
                        },
                    },
                    "transformer": {
                        "type": "object",
                        "properties": {
                            "steps": {"type": "array", "items": {"type": "string"}},
                            "options": {"type": "object"},
                            "column_map": {"type": "object"},
                        },
                    },
                    "loader": {
                        "type": "object",
                        "required": ["target_table", "mode"],
                        "properties": {
                            "target_table": {"type": "string"},
                            "mode": {"enum": ["full_reload", "idempotent_period", "append"]},
                            "period_column": {"type": "string"},
                            "key_columns": {"type": "array", "items": {"type": "string"}},
                        },
                    },
                },
            },
        },
    },
}
