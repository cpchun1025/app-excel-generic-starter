"""Service boundary for explicitly fake or real execution bundles."""

from typing import Any, Protocol

from excel_automation.config.models import Job, Step
from excel_automation.runner.context import RunContext


class ExecutionServices(Protocol):
    """Operations an orchestrator requires from a runtime adapter bundle."""

    fake_execution: bool

    def preflight(self, job: Job) -> None: ...

    def stage(self, job: Job, context: RunContext) -> None: ...

    def execute(self, step: Step, context: RunContext) -> dict[str, Any]: ...


class FakeExecutionServices:
    """In-memory simulation; it never reads, edits, or publishes workbooks."""

    fake_execution = True

    def __init__(self, *, fail_step: str | None = None) -> None:
        self.events: list[tuple[str, str]] = []
        self.fail_step = fail_step
        self._staged_output_token = None

    def preflight(self, job: Job) -> None:
        """Fake runs validate orchestration only, never inspect real workbooks."""

    def stage(self, job: Job, context: RunContext) -> None:
        self.events.append(("stage", job.id))
        self._staged_output_token = context.staged_output

    def execute(self, step: Step, context: RunContext) -> dict[str, Any]:
        self.events.append((step.type, step.id))
        if step.id == self.fail_step:
            raise RuntimeError("Injected fake-service failure")

        if step.type == "validate":
            if self._staged_output_token != context.staged_output:
                raise RuntimeError("Fake validation has no staged-output token")
            results = []
            for check in step.definition["checks"]:
                if check["type"] == "file_exists":
                    passed = self._staged_output_token == context.staged_output
                    if not passed:
                        raise RuntimeError("Fake staged output is unavailable")
                    results.append({"type": "file_exists", "simulated_pass": True})
                else:
                    results.append({"type": check["type"], "simulated_pass": True})
            return {
                "simulated_checks": results,
                "simulated": True,
            }
        if step.type == "email":
            return {
                "prepared": bool(context.job.email and context.job.email["enabled"]),
                "sent": False,
                "simulated": True,
            }
        return {"simulated": True}
