from pathlib import Path
import logging
from dataclasses import replace

import pytest

from excel_automation.cli import (
    EXIT_CONFIGURATION,
    EXIT_OK,
    EXIT_RUNTIME,
    main,
)
from excel_automation.config.loader import load_config
from excel_automation.config.validation import load_and_validate, validate_document
from excel_automation.errors import (
    ConfigurationError,
    OutputCollisionError,
)
from excel_automation.plugins.example_allowlist import APPROVED_EXAMPLE_PLUGIN_IDS
from excel_automation.runner.orchestrator import run_job
from excel_automation.runner.services import FakeExecutionServices


ROOT = Path(__file__).resolve().parents[2]
EXAMPLE_CONFIG = ROOT / "config" / "jobs.example.yaml"
def test_example_configuration_schema_validates_all_workflows():
    config = load_and_validate(EXAMPLE_CONFIG, plugin_ids=APPROVED_EXAMPLE_PLUGIN_IDS)

    assert set(config.jobs) == {
        "fake-python-only",
        "fake-macro-only",
        "fake-python-then-macro",
    }
    assert [step.type for step in config.jobs["fake-python-only"].steps] == [
        "python",
        "validate",
    ]
    assert [step.type for step in config.jobs["fake-macro-only"].steps] == [
        "copy",
        "macro",
        "validate",
    ]
    assert [
        step.type for step in config.jobs["fake-python-then-macro"].steps
    ] == ["copy", "python", "macro", "validate", "email"]


def test_unknown_plugin_id_fails_before_execution(tmp_path):
    document, source = load_config(EXAMPLE_CONFIG)
    document["jobs"]["fake-python-only"]["steps"][0]["plugin"] = "untrusted.run"

    with pytest.raises(ConfigurationError, match="not registered"):
        validate_document(document, source, plugin_ids=APPROVED_EXAMPLE_PLUGIN_IDS)


def test_strict_schema_rejects_unknown_field(tmp_path):
    document, source = load_config(EXAMPLE_CONFIG)
    document["jobs"]["fake-python-only"]["unexpected"] = "ignored?"

    with pytest.raises(ConfigurationError, match="unknown field"):
        validate_document(document, source, plugin_ids=APPROVED_EXAMPLE_PLUGIN_IDS)


def test_schema_version_unknown_step_duplicate_ids_and_missing_input_fail():
    invalidations = []

    document, source = load_config(EXAMPLE_CONFIG)
    document["schema_version"] = 2
    invalidations.append(document)

    document, source = load_config(EXAMPLE_CONFIG)
    document["jobs"]["fake-python-only"]["steps"][0]["type"] = "shell"
    invalidations.append(document)

    document, source = load_config(EXAMPLE_CONFIG)
    steps = document["jobs"]["fake-python-only"]["steps"]
    steps[1]["id"] = steps[0]["id"]
    invalidations.append(document)

    document, source = load_config(EXAMPLE_CONFIG)
    document["jobs"]["fake-python-only"]["workbook"]["base"] = "missing"
    invalidations.append(document)

    for invalid in invalidations:
        with pytest.raises(ConfigurationError):
            validate_document(invalid, source, plugin_ids=APPROVED_EXAMPLE_PLUGIN_IDS)


def test_output_must_not_resolve_to_any_input():
    document, source = load_config(EXAMPLE_CONFIG)
    job = document["jobs"]["fake-python-only"]
    job["workbook"]["output"] = job["inputs"]["source"]

    with pytest.raises(ConfigurationError, match="differ from every input"):
        validate_document(document, source, plugin_ids=APPROVED_EXAMPLE_PLUGIN_IDS)


def test_macro_is_not_compatible_with_direct_adapter(tmp_path):
    document, source = load_config(EXAMPLE_CONFIG)
    job = document["jobs"]["fake-macro-only"]
    job["workbook"]["adapter"] = "direct_xlsx"
    job["inputs"]["template"] = "../examples/workbooks/fake_month_end.xlsx"
    job["workbook"]["output"] = "../examples/output/fake_month_end_result.xlsx"

    with pytest.raises(ConfigurationError, match="requires excel_desktop"):
        validate_document(document, source, plugin_ids=APPROVED_EXAMPLE_PLUGIN_IDS)


def test_enabled_email_requires_validate_before_final_email(tmp_path):
    document, source = load_config(EXAMPLE_CONFIG)
    steps = document["jobs"]["fake-python-then-macro"]["steps"]
    steps.remove(next(step for step in steps if step["type"] == "validate"))

    with pytest.raises(ConfigurationError, match="requires validation"):
        validate_document(document, source, plugin_ids=APPROVED_EXAMPLE_PLUGIN_IDS)


def test_yaml_duplicate_keys_are_rejected(tmp_path):
    path = tmp_path / "duplicate.yaml"
    path.write_text("schema_version: 1\nschema_version: 1\njobs: {}\n", encoding="utf-8")

    with pytest.raises(ConfigurationError, match="Invalid YAML"):
        load_config(path)


def test_fake_services_execute_steps_in_order_without_publishing(tmp_path):
    config = load_and_validate(EXAMPLE_CONFIG, plugin_ids=APPROVED_EXAMPLE_PLUGIN_IDS)
    job = config.jobs["fake-python-then-macro"]
    services = FakeExecutionServices()

    result = run_job(job, services, dry_run=True)

    assert services.events == [
        ("stage", job.id),
        ("copy", "copy-forecast-inputs"),
        ("python", "prepare-forecast"),
        ("macro", "calculate-forecast"),
        ("validate", "validate-forecast-output"),
        ("email", "prepare-forecast-email"),
    ]
    assert result["status"] == "fake_simulation_succeeded"
    assert result["fake_execution"] is True
    assert result["dry_run"] is True
    assert result["output_published"] is False
    assert not job.output.exists()


def test_fake_file_exists_check_uses_virtual_staged_output_not_final_path():
    config = load_and_validate(EXAMPLE_CONFIG, plugin_ids=APPROVED_EXAMPLE_PLUGIN_IDS)
    job = config.jobs["fake-python-only"]

    class InspectStagedOutput(FakeExecutionServices):
        def execute(self, step, context):
            if step.type == "validate":
                assert context.staged_output.parent == context.workspace
                assert not context.staged_output.exists()
                assert not context.job.output.exists()
            return super().execute(step, context)

    result = run_job(job, InspectStagedOutput())
    assert result["status"] == "fake_simulation_succeeded"


@pytest.mark.parametrize("dry_run", [False, True])
def test_fake_execution_never_publishes_output(tmp_path, dry_run):
    config = load_and_validate(EXAMPLE_CONFIG, plugin_ids=APPROVED_EXAMPLE_PLUGIN_IDS)
    job = config.jobs["fake-python-only"]

    result = run_job(job, FakeExecutionServices(), dry_run=dry_run)

    assert result["output_published"] is False
    assert not job.output.exists()


def test_output_collision_is_rejected_for_fake_and_dry_runs(tmp_path):
    config = load_and_validate(EXAMPLE_CONFIG, plugin_ids=APPROVED_EXAMPLE_PLUGIN_IDS)
    job = replace(config.jobs["fake-python-only"], output=tmp_path / "output" / "result.xlsx")
    job.output.parent.mkdir(parents=True, exist_ok=True)
    job.output.write_text("existing output", encoding="utf-8")
    services = FakeExecutionServices()
    try:
        for dry_run in (False, True):
            with pytest.raises(OutputCollisionError):
                run_job(job, services, dry_run=dry_run)
        assert services.events == []
    finally:
        job.output.unlink()


def test_failed_fake_step_stops_following_steps(tmp_path):
    config = load_and_validate(EXAMPLE_CONFIG, plugin_ids=APPROVED_EXAMPLE_PLUGIN_IDS)
    job = config.jobs["fake-python-then-macro"]
    services = FakeExecutionServices(fail_step="calculate-forecast")

    with pytest.raises(RuntimeError, match="Injected fake-service failure"):
        run_job(job, services)

    assert [event[1] for event in services.events] == [
        job.id,
        "copy-forecast-inputs",
        "prepare-forecast",
        "calculate-forecast",
    ]
    assert not job.output.exists()


def test_failed_fake_validation_prevents_email_step():
    config = load_and_validate(EXAMPLE_CONFIG, plugin_ids=APPROVED_EXAMPLE_PLUGIN_IDS)
    job = config.jobs["fake-python-then-macro"]
    services = FakeExecutionServices(fail_step="validate-forecast-output")

    with pytest.raises(RuntimeError, match="Injected fake-service failure"):
        run_job(job, services)

    assert all(event[1] != "prepare-forecast-email" for event in services.events)


def test_cli_refuses_unimplemented_real_steps_and_labels_fake_mode(capsys):
    assert main(["--config", str(EXAMPLE_CONFIG), "run", "fake-python-only"]) == EXIT_RUNTIME
    assert "Unimplemented real step(s): python" in capsys.readouterr().err

    assert main(
        [
            "--config",
            str(EXAMPLE_CONFIG),
            "run",
            "fake-python-only",
            "--fake-services",
            "--dry-run",
        ]
    ) == EXIT_OK
    output = capsys.readouterr().out
    assert "FAKE SIMULATION SUCCEEDED" in output
    assert "no Excel workbook was read or modified" in output
    assert "no output was published" in output


def test_cli_validate_and_list_jobs_distinguish_production_status(capsys):
    assert main(["--config", str(EXAMPLE_CONFIG), "validate"]) == EXIT_OK
    assert "Validated 3 job(s)" in capsys.readouterr().out

    assert main(["--config", str(EXAMPLE_CONFIG), "list-jobs"]) == EXIT_OK
    output = capsys.readouterr().out
    assert "schema-valid; fake-executable; production-executable=no" in output


def test_cli_unknown_job_is_configuration_error(capsys):
    assert main(
        [
            "--config",
            str(EXAMPLE_CONFIG),
            "run",
            "missing-job",
            "--fake-services",
        ]
    ) == EXIT_CONFIGURATION
    assert "Unknown job ID" in capsys.readouterr().err


def test_step_failure_log_omits_exception_payload(caplog):
    config = load_and_validate(EXAMPLE_CONFIG, plugin_ids=APPROVED_EXAMPLE_PLUGIN_IDS)
    job = config.jobs["fake-python-only"]
    sensitive_text = "sensitive-cell-value-and-secret-token"

    class SensitiveFailureServices(FakeExecutionServices):
        def execute(self, step, context):
            self.events.append((step.type, step.id))
            raise RuntimeError(sensitive_text)

    test_logger = logging.getLogger("phase1_test")
    with caplog.at_level(logging.INFO, logger="phase1_test"):
        with pytest.raises(RuntimeError, match="sensitive-cell-value"):
            run_job(
                job,
                SensitiveFailureServices(),
                logger=test_logger,
            )

    assert sensitive_text not in caplog.text
    assert "status=failed" in caplog.text
