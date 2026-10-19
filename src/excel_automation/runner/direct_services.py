"""Real execution bundle limited to staging and Phase 2 validation."""

from typing import Any

from excel_automation.config.models import Job, Step
from excel_automation.errors import ExecutionError
from excel_automation.runner.context import RunContext
from excel_automation.workbook.copies import stage_copy
from excel_automation.workbook.direct_xlsx import DirectXlsx
from excel_automation.workbook.validation import check_definition, validate_output


class DirectXlsxServices:
    fake_execution = False

    @staticmethod
    def capability_error(job: Job) -> str | None:
        if job.adapter != "direct_xlsx":
            return "Real execution requires direct_xlsx; desktop Excel is not implemented"
        unsupported = [step.type for step in job.steps if step.type != "validate"]
        if unsupported:
            return (
                "Unimplemented real step(s): " + ", ".join(unsupported)
                + "; Phase 2 supports staging and validate steps only"
            )
        if job.email and job.email["enabled"]:
            return "Real email/rendering execution is not implemented"
        for step in job.steps:
            try:
                for check in step.definition["checks"]:
                    check_definition(check)
            except ExecutionError as exc:
                return str(exc)
        return None

    def preflight(self, job: Job) -> None:
        reason = self.capability_error(job)
        if reason:
            raise ExecutionError(reason)
        for path in job.inputs.values():
            with DirectXlsx(path):
                pass

    def stage(self, job: Job, context: RunContext) -> None:
        stage_copy(job.inputs[job.base], context.staged_output)
        # Recheck the actual staged bytes, not just the preflight source.
        with DirectXlsx(context.staged_output):
            pass

    def execute(self, step: Step, context: RunContext) -> dict[str, Any]:
        if step.type != "validate":
            raise ExecutionError("Unimplemented step reached direct_xlsx runtime")
        validate_output(context.staged_output, step.definition["checks"])
        return {"simulated": False}
