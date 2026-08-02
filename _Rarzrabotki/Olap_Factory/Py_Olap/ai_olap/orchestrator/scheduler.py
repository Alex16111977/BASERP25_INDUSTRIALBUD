"""EtlScheduler — APScheduler BlockingScheduler with cron triggers from pipelines/."""
from __future__ import annotations

from pathlib import Path

import structlog
from apscheduler.schedulers.blocking import BlockingScheduler
from apscheduler.triggers.cron import CronTrigger

from ..config.loader import discover_pipelines, load_pipeline_config
from ..config.validator import validate_pipeline
from .runner import run_pipeline

log = structlog.get_logger().bind(component="scheduler")


class EtlScheduler:
    def __init__(self, pipelines_dir: str | Path = "pipelines") -> None:
        self.pipelines_dir = Path(pipelines_dir)
        self.scheduler = BlockingScheduler(
            job_defaults={"coalesce": True, "max_instances": 1, "misfire_grace_time": 600}
        )

    def register_jobs(self) -> int:
        registered = 0
        for path in discover_pipelines(self.pipelines_dir):
            cfg = load_pipeline_config(path)
            validate_pipeline(cfg, source=path.name)
            cron = cfg.get("schedule")
            if not cron:
                log.info("skip no schedule", pipeline=cfg["pipeline_id"])
                continue
            trigger = CronTrigger.from_crontab(cron)
            self.scheduler.add_job(
                run_pipeline,
                trigger=trigger,
                kwargs={"cfg": cfg, "script": path.stem},
                id=cfg["pipeline_id"],
                name=cfg["pipeline_id"],
                replace_existing=True,
            )
            registered += 1
            log.info("job registered", pipeline=cfg["pipeline_id"], cron=cron)
        return registered

    def list_jobs(self) -> list[str]:
        return [j.id for j in self.scheduler.get_jobs()]

    def start(self) -> None:
        log.info("scheduler starting", jobs=self.list_jobs())
        self.scheduler.start()

    def shutdown(self, wait: bool = False) -> None:
        self.scheduler.shutdown(wait=wait)
