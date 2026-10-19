"""Temporary-workspace orchestration for injected execution services."""

import logging
import os
import sys
import tempfile
import uuid
from collections.abc import Callable
from pathlib import Path
from typing import Any

from excel_automation.config.models import Job
from excel_automation.errors import OutputCollisionError
from excel_automation.runner.context import RunContext
from excel_automation.runner.services import ExecutionServices
from excel_automation.workbook.copies import publish_output


def run_job(
    job: Job,
    services: ExecutionServices,
    *,
    dry_run: bool = False,
    on_event: Callable[[str], None] | None = None,
    logger: logging.Logger | None = None,
) -> dict[str, Any]:
    """Preflight every step, execute in isolation, and publish real runs only."""
    services.preflight(job)
    if os.path.lexists(job.output):
        raise OutputCollisionError(f"Output path already exists: {job.output}")
    run_id = uuid.uuid4().hex
    log = logger or logging.getLogger("excel_automation")

    real_publish = not services.fake_execution and not dry_run
    if real_publish:
        job.output.parent.mkdir(parents=True, exist_ok=True)
    workspace_parent = job.output.parent if real_publish else None
    with tempfile.TemporaryDirectory(
        prefix=f"excel-job-{job.id}-", dir=workspace_parent
    ) as workspace_name:
        workspace = Path(workspace_name)
        staged_output = workspace / ("staged-output" + job.output.suffix)
        context = RunContext(
            run_id=run_id,
            job=job,
            workspace=workspace,
            staged_output=staged_output,
            dry_run=dry_run,
            fake_execution=services.fake_execution,
        )
        mode = "fake" if services.fake_execution else "direct_xlsx"
        log.info("run_id=%s job=%s status=started mode=%s dry_run=%s", run_id, job.id, mode, dry_run)
        if on_event:
            label = "fake simulation; no Excel processing" if services.fake_execution else "real direct_xlsx"
            on_event(f"Run {run_id}: {job.id} ({label})")

        current_step = "stage"
        completed = False
        try:
            services.stage(job, context)
            log.info("run_id=%s job=%s step=stage status=success mode=%s", run_id, job.id, mode)
            for step in job.steps:
                current_step = step.id
                log.info(
                    "run_id=%s job=%s step=%s status=started mode=%s",
                    run_id,
                    job.id,
                    step.id,
                    mode,
                )
                services.execute(step, context)
                log.info(
                    "run_id=%s job=%s step=%s status=success mode=%s",
                    run_id,
                    job.id,
                    step.id,
                    mode,
                )
                if on_event:
                    verb = "simulated" if services.fake_execution else "completed"
                    on_event(f"Step {step.id}: {verb} {step.type}")
            if real_publish:
                current_step = "publish"
                publish_output(context.staged_output, job.output)
            completed = True
            log.info(
                "run_id=%s job=%s status=success mode=%s output_published=%s",
                run_id,
                job.id,
                mode,
                real_publish,
            )
            return {
                "run_id": run_id,
                "job_id": job.id,
                "status": "fake_simulation_succeeded" if services.fake_execution else "real_execution_succeeded",
                "fake_execution": services.fake_execution,
                "dry_run": dry_run,
                "output_published": real_publish,
                "steps_completed": len(job.steps),
            }
        finally:
            exception_type = sys.exc_info()[0]
            log.info(
                "run_id=%s job=%s step=%s status=%s error_type=%s",
                run_id,
                job.id,
                current_step,
                "success" if completed else "failed",
                exception_type.__name__ if exception_type else "none",
            )
