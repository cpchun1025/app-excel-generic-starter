"""Execution metadata shared with explicitly injected services."""

from dataclasses import dataclass
from pathlib import Path

from excel_automation.config.models import Job


@dataclass(frozen=True)
class RunContext:
    run_id: str
    job: Job
    workspace: Path
    staged_output: Path
    dry_run: bool
    fake_execution: bool
