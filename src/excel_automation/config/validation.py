"""Strict validation for configuration schema version 1."""

import re
from collections.abc import Mapping, Set
from pathlib import Path
from typing import Any

from excel_automation.config.models import Configuration, Job, Step
from excel_automation.errors import ConfigurationError

JOB_ID = re.compile(r"^[a-z0-9][a-z0-9_-]*$")
PLUGIN_ID = re.compile(r"^[a-z][a-z0-9_]*(?:\.[a-z][a-z0-9_]*)+$")
ENV_REF = re.compile(r"^\$\{[A-Za-z_][A-Za-z0-9_]*\}$")
RANGE = re.compile(r"^\$?[A-Z]{1,3}\$?[1-9][0-9]*(?::\$?[A-Z]{1,3}\$?[1-9][0-9]*)?$")
SCALAR_TYPES = (str, int, float, bool)

TOP_KEYS = {"schema_version", "jobs"}
JOB_KEYS = {"inputs", "workbook", "steps", "email"}
WORKBOOK_KEYS = {"base", "output", "adapter"}
STEP_COMMON = {"id", "type"}
STEP_KEYS = {
    "copy": STEP_COMMON | {"source", "destination", "mode"},
    "python": STEP_COMMON | {"plugin", "options"},
    "macro": STEP_COMMON | {"workbook", "name", "args"},
    "validate": STEP_COMMON | {"checks"},
    "email": STEP_COMMON,
}
CHECK_KEYS = {
    "file_exists": {"type"},
    "sheet_exists": {"type", "sheet"},
    "range_nonempty": {"type", "sheet", "range"},
    "range_equals": {"type", "sheet", "range", "expected"},
}
EMAIL_KEYS = {
    "enabled", "send", "smtp", "subject", "body", "delivery_key", "outputs",
}
SMTP_KEYS = {"host", "port", "tls", "sender", "recipients", "username", "password"}
EMAIL_OUTPUT_KEYS = {"type", "sheet", "range", "placement"}
KNOWN_STEP_TYPES = set(STEP_KEYS)


def validate_document(
    document: Mapping[str, Any],
    source_path: Path,
    *,
    plugin_ids: Set[str] = frozenset(),
) -> Configuration:
    _keys(document, TOP_KEYS, "configuration")
    if type(document.get("schema_version")) is not int or document["schema_version"] != 1:
        raise ConfigurationError("schema_version must be integer 1")
    raw_jobs = _mapping(document.get("jobs"), "jobs")
    if not raw_jobs:
        raise ConfigurationError("jobs must contain at least one job")

    jobs: dict[str, Job] = {}
    for job_id, raw_job in raw_jobs.items():
        if not isinstance(job_id, str) or not JOB_ID.fullmatch(job_id):
            raise ConfigurationError(f"Invalid job ID: {job_id!r}")
        if job_id in jobs:
            raise ConfigurationError(f"Duplicate job ID: {job_id}")
        jobs[job_id] = _validate_job(
            job_id, _mapping(raw_job, f"jobs.{job_id}"), source_path, plugin_ids
        )
    return Configuration(source_path=source_path, jobs=jobs)


def _validate_job(
    job_id: str,
    raw: Mapping[str, Any],
    source_path: Path,
    plugin_ids: Set[str],
) -> Job:
    _keys(raw, JOB_KEYS, f"jobs.{job_id}")
    raw_inputs = _mapping(raw.get("inputs"), f"jobs.{job_id}.inputs")
    if not raw_inputs:
        raise ConfigurationError(f"jobs.{job_id}.inputs must not be empty")
    inputs: dict[str, Path] = {}
    for alias, value in raw_inputs.items():
        if not isinstance(alias, str) or not JOB_ID.fullmatch(alias):
            raise ConfigurationError(f"Invalid input alias in job {job_id}: {alias!r}")
        inputs[alias] = _path(value, source_path.parent, f"jobs.{job_id}.inputs.{alias}")

    workbook = _mapping(raw.get("workbook"), f"jobs.{job_id}.workbook")
    _keys(workbook, WORKBOOK_KEYS, f"jobs.{job_id}.workbook")
    base = _string(workbook.get("base"), f"jobs.{job_id}.workbook.base")
    if base not in inputs:
        raise ConfigurationError(f"jobs.{job_id}.workbook.base must name an input alias")
    output = _path(workbook.get("output"), source_path.parent, f"jobs.{job_id}.workbook.output")
    adapter = _string(workbook.get("adapter"), f"jobs.{job_id}.workbook.adapter")
    if adapter not in {"direct_xlsx", "excel_desktop"}:
        raise ConfigurationError(f"jobs.{job_id}.workbook.adapter is unsupported: {adapter}")
    if output in inputs.values():
        raise ConfigurationError(f"jobs.{job_id}.workbook.output must differ from every input")
    if adapter == "direct_xlsx":
        direct_paths = [*inputs.values(), output]
        if any(path.suffix.lower() != ".xlsx" for path in direct_paths):
            raise ConfigurationError(
                f"jobs.{job_id} direct_xlsx adapter accepts only .xlsx paths"
            )

    raw_steps = raw.get("steps")
    if not isinstance(raw_steps, list) or not raw_steps:
        raise ConfigurationError(f"jobs.{job_id}.steps must be a non-empty list")
    steps: list[Step] = []
    seen_ids: set[str] = set()
    for index, raw_step in enumerate(raw_steps):
        step = _mapping(raw_step, f"jobs.{job_id}.steps[{index}]")
        step_id, step_type = _validate_step(
            step, job_id, index, inputs, adapter, plugin_ids
        )
        if step_id in seen_ids:
            raise ConfigurationError(f"Duplicate step ID in {job_id}: {step_id}")
        seen_ids.add(step_id)
        steps.append(Step(step_id, step_type, dict(step)))

    email = _validate_email(raw.get("email"), job_id)
    email_positions = [i for i, step in enumerate(steps) if step.type == "email"]
    if email is None or email["enabled"] is False:
        if email_positions:
            raise ConfigurationError(f"jobs.{job_id} has an email step but email is disabled")
    else:
        if len(email_positions) != 1 or email_positions[0] != len(steps) - 1:
            raise ConfigurationError(
                f"jobs.{job_id} requires exactly one final email step when email is enabled"
            )
        validation_positions = [i for i, step in enumerate(steps) if step.type == "validate"]
        if not validation_positions or max(validation_positions) >= email_positions[0]:
            raise ConfigurationError(f"jobs.{job_id} requires validation before email")

    return Job(job_id, inputs, base, output, adapter, tuple(steps), email)


def _validate_step(
    step: Mapping[str, Any],
    job_id: str,
    index: int,
    inputs: Mapping[str, Path],
    adapter: str,
    plugin_ids: Set[str],
) -> tuple[str, str]:
    where = f"jobs.{job_id}.steps[{index}]"
    step_id = _string(step.get("id"), f"{where}.id")
    if not JOB_ID.fullmatch(step_id):
        raise ConfigurationError(f"{where}.id is invalid")
    step_type = _string(step.get("type"), f"{where}.type")
    if step_type not in STEP_KEYS:
        raise ConfigurationError(f"{where}.type is unsupported: {step_type}")
    _keys(step, STEP_KEYS[step_type], where)

    if step_type == "copy":
        source = _mapping(step.get("source"), f"{where}.source")
        destination = _mapping(step.get("destination"), f"{where}.destination")
        _keys(source, {"workbook", "sheet", "range"}, f"{where}.source")
        _keys(destination, {"workbook", "sheet", "range"}, f"{where}.destination")
        alias = _string(source.get("workbook"), f"{where}.source.workbook")
        if alias not in inputs:
            raise ConfigurationError(f"{where}.source.workbook is not an input alias")
        if _string(destination.get("workbook"), f"{where}.destination.workbook") != "output":
            raise ConfigurationError(f"{where}.destination.workbook must be 'output'")
        _string(source.get("sheet"), f"{where}.source.sheet")
        _string(destination.get("sheet"), f"{where}.destination.sheet")
        mode = _string(step.get("mode"), f"{where}.mode")
        if mode not in {"values", "values_and_formatting", "worksheet"}:
            raise ConfigurationError(f"{where}.mode is unsupported")
        if mode == "worksheet":
            if "range" in source or "range" in destination:
                raise ConfigurationError(f"{where} worksheet mode cannot include ranges")
        else:
            _range(source.get("range"), f"{where}.source.range")
            _range(destination.get("range"), f"{where}.destination.range")
    elif step_type == "python":
        plugin = _string(step.get("plugin"), f"{where}.plugin")
        if not PLUGIN_ID.fullmatch(plugin):
            raise ConfigurationError(f"{where}.plugin must be a qualified approved plugin ID")
        if plugin not in plugin_ids:
            raise ConfigurationError(f"{where}.plugin is not registered: {plugin}")
        options = step.get("options", {})
        _mapping(options, f"{where}.options")
    elif step_type == "macro":
        if adapter != "excel_desktop":
            raise ConfigurationError(f"{where} macro requires excel_desktop adapter")
        if _string(step.get("workbook"), f"{where}.workbook") != "output":
            raise ConfigurationError(f"{where}.workbook must be 'output'")
        _string(step.get("name"), f"{where}.name")
        args = step.get("args", [])
        if not isinstance(args, list) or any(type(arg) not in SCALAR_TYPES for arg in args):
            raise ConfigurationError(f"{where}.args must be a list of scalar values")
    elif step_type == "validate":
        checks = step.get("checks")
        if not isinstance(checks, list) or not checks:
            raise ConfigurationError(f"{where}.checks must be a non-empty list")
        for check_index, raw_check in enumerate(checks):
            check_where = f"{where}.checks[{check_index}]"
            check = _mapping(raw_check, check_where)
            check_type = _string(check.get("type"), f"{check_where}.type")
            if check_type not in CHECK_KEYS:
                raise ConfigurationError(f"{check_where}.type is unsupported: {check_type}")
            _keys(check, CHECK_KEYS[check_type], check_where)
            if check_type in {"sheet_exists", "range_nonempty", "range_equals"}:
                _string(check.get("sheet"), f"{check_where}.sheet")
            if check_type in {"range_nonempty", "range_equals"}:
                _range(check.get("range"), f"{check_where}.range")
            if check_type == "range_equals" and (
                "expected" not in check or not _valid_expected(check["expected"])
            ):
                raise ConfigurationError(f"{check_where}.expected must be a scalar or matrix")
    return step_id, step_type


def _validate_email(raw: Any, job_id: str) -> dict[str, Any] | None:
    if raw is None:
        return None
    email = _mapping(raw, f"jobs.{job_id}.email")
    _keys(email, EMAIL_KEYS, f"jobs.{job_id}.email")
    enabled = _boolean(email.get("enabled"), f"jobs.{job_id}.email.enabled")
    send = email.get("send", True)
    if type(send) is not bool:
        raise ConfigurationError(f"jobs.{job_id}.email.send must be a boolean")
    result = dict(email)
    result["enabled"] = enabled
    result["send"] = send
    if not enabled:
        if set(email) - {"enabled", "send"}:
            raise ConfigurationError(f"jobs.{job_id}.email disabled cannot include delivery settings")
        return result

    smtp = _mapping(email.get("smtp"), f"jobs.{job_id}.email.smtp")
    _keys(smtp, SMTP_KEYS, f"jobs.{job_id}.email.smtp")
    for key in ("host", "sender"):
        _string(smtp.get(key), f"jobs.{job_id}.email.smtp.{key}")
    port = smtp.get("port")
    if type(port) is not int or not 1 <= port <= 65535:
        raise ConfigurationError(f"jobs.{job_id}.email.smtp.port must be between 1 and 65535")
    tls = smtp.get("tls")
    if not isinstance(tls, str) or tls not in {"none", "starttls", "implicit"}:
        raise ConfigurationError(f"jobs.{job_id}.email.smtp.tls is invalid")
    recipients = smtp.get("recipients")
    if not isinstance(recipients, list) or not recipients:
        raise ConfigurationError(f"jobs.{job_id}.email.smtp.recipients must be non-empty")
    for recipient in recipients:
        _string(recipient, f"jobs.{job_id}.email.smtp.recipients[]")
    for credential in ("username", "password"):
        value = smtp.get(credential)
        if value is not None and (not isinstance(value, str) or not ENV_REF.fullmatch(value)):
            raise ConfigurationError(
                f"jobs.{job_id}.email.smtp.{credential} must be an environment reference"
            )
    if ("username" in smtp) != ("password" in smtp):
        raise ConfigurationError("SMTP username and password references must be provided together")
    for key in ("subject", "body", "delivery_key"):
        _string(email.get(key), f"jobs.{job_id}.email.{key}")
    outputs = email.get("outputs")
    if not isinstance(outputs, list) or not outputs:
        raise ConfigurationError(f"jobs.{job_id}.email.outputs must be non-empty")
    for index, raw_output in enumerate(outputs):
        where = f"jobs.{job_id}.email.outputs[{index}]"
        output = _mapping(raw_output, where)
        _keys(output, EMAIL_OUTPUT_KEYS, where)
        output_type = _string(output.get("type"), f"{where}.type")
        placement = _string(output.get("placement"), f"{where}.placement")
        if output_type not in {"html_table", "range_image", "html_attachment", "workbook"}:
            raise ConfigurationError(f"{where}.type is unsupported: {output_type}")
        expected_placement = "body" if output_type == "html_table" else "attachment"
        if placement != expected_placement:
            raise ConfigurationError(f"{where}.placement must be {expected_placement}")
        if output_type != "workbook":
            _string(output.get("sheet"), f"{where}.sheet")
            _range(output.get("range"), f"{where}.range")
    return result


def _mapping(value: Any, where: str) -> Mapping[str, Any]:
    if not isinstance(value, Mapping) or any(not isinstance(key, str) for key in value):
        raise ConfigurationError(f"{where} must be a string-keyed mapping")
    return value


def _keys(value: Mapping[str, Any], allowed: set[str], where: str) -> None:
    unknown = set(value) - allowed
    if unknown:
        raise ConfigurationError(f"{where} contains unknown field(s): {', '.join(sorted(unknown))}")


def _string(value: Any, where: str) -> str:
    if not isinstance(value, str) or not value.strip():
        raise ConfigurationError(f"{where} must be a non-empty string")
    return value


def _boolean(value: Any, where: str) -> bool:
    if type(value) is not bool:
        raise ConfigurationError(f"{where} must be a boolean")
    return value


def _path(value: Any, base: Path, where: str) -> Path:
    text = _string(value, where)
    path = Path(text).expanduser()
    return (base / path).resolve() if not path.is_absolute() else path.resolve()


def _range(value: Any, where: str) -> str:
    text = _string(value, where)
    if not RANGE.fullmatch(text):
        raise ConfigurationError(f"{where} must be a simple A1-style cell or range")
    return text


def _valid_expected(value: Any) -> bool:
    if type(value) in SCALAR_TYPES:
        return True
    if isinstance(value, list):
        return all(type(item) in SCALAR_TYPES for item in value) or all(
            isinstance(row, list) and all(type(item) in SCALAR_TYPES for item in row)
            for row in value
        )
    return False


def load_and_validate(
    path: Path, *, plugin_ids: Set[str] = frozenset()
) -> Configuration:
    from excel_automation.config.loader import load_config

    document, config_path = load_config(path)
    return validate_document(document, config_path, plugin_ids=plugin_ids)
