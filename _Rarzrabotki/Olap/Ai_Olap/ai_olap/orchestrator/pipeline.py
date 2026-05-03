"""Pipeline executes Extract -> Transform -> Load for a single step config.

A pipeline file may contain N steps; each is run sequentially by Pipeline.run().
"""
from __future__ import annotations

import datetime as dt

import structlog

from ..core.exceptions import ETLException
from ..extractors.factory import make_extractor
from ..loaders.factory import make_loader
from ..transformers import pipeline as tpipe

log = structlog.get_logger().bind(component="pipeline")


class StepResult:
    def __init__(self, step_id: str, rows_extracted: int, rows_loaded: int) -> None:
        self.step_id = step_id
        self.rows_extracted = rows_extracted
        self.rows_loaded = rows_loaded


class Pipeline:
    def __init__(self, config: dict, *, period: dt.date | None = None) -> None:
        self.config = config
        self.pipeline_id = config["pipeline_id"]
        self.period = period

    def run(self) -> list[StepResult]:
        results: list[StepResult] = []
        for step in self.config["steps"]:
            sid = step["step_id"]
            log.info("step start", pipeline=self.pipeline_id, step=sid)

            extractor_cfg = dict(step["extractor"])
            if extractor_cfg.get("auto_period_params") and self.period:
                p = self.period
                next_month = dt.date(p.year + (p.month == 12), (p.month % 12) + 1, 1)
                # BaseERP cluster backend stores _Date_Time with +2000 year offset.
                offset = int(extractor_cfg.get("period_offset_years", 2000))
                p_db = dt.date(p.year + offset, p.month, p.day)
                n_db = dt.date(next_month.year + offset, next_month.month, next_month.day)
                extractor_cfg["params"] = [p_db, n_db]
            extractor = make_extractor(extractor_cfg)
            rows = extractor.extract()

            xform = step.get("transformer") or {}
            steps_list = xform.get("steps") or []
            options = xform.get("options") or {}
            if "column_mapper" in steps_list and "column_mapper" not in options:
                if xform.get("column_map"):
                    options["column_mapper"] = {"column_map": xform["column_map"]}
            if steps_list:
                rows = tpipe.apply(rows, steps_list, options)

            loader_cfg = dict(step["loader"])
            if loader_cfg.get("mode") == "idempotent_period" and self.period:
                loader_cfg["period"] = self.period.isoformat()
            loader = make_loader(loader_cfg)
            inserted = loader.load(rows)

            log.info("step done", pipeline=self.pipeline_id, step=sid,
                     extracted=len(rows), loaded=inserted)
            results.append(StepResult(sid, len(rows), inserted))
        return results
