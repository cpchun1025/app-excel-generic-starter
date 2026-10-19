"""Configuration inspection, explicit simulation, and constrained real execution."""

import argparse
import logging
import sys
from pathlib import Path

from excel_automation.config.models import Configuration
from excel_automation.config.validation import load_and_validate
from excel_automation.errors import AutomationError, ConfigurationError
from excel_automation.logging_config import configure_logging
from excel_automation.plugins.example_allowlist import APPROVED_EXAMPLE_PLUGIN_IDS
from excel_automation.runner.direct_services import DirectXlsxServices
from excel_automation.runner.orchestrator import run_job
from excel_automation.runner.services import FakeExecutionServices
from excel_automation.samples import generate_sample

EXIT_OK = 0
EXIT_CONFIGURATION = 2
EXIT_RUNTIME = 3


def _parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="excel-automation",
        description="Validate, simulate, or run supported direct_xlsx jobs.",
    )
    parser.add_argument(
        "--config",
        type=Path,
        default=Path("config/jobs.example.yaml"),
        help="YAML configuration path (default: config/jobs.example.yaml)",
    )
    subparsers = parser.add_subparsers(dest="command", required=True)
    subparsers.add_parser("list-jobs", help="List schema-valid configured jobs")
    subparsers.add_parser("validate", help="Validate configuration without running jobs")
    sample_parser = subparsers.add_parser("generate-samples", help="Generate a plain, non-sensitive .xlsx")
    sample_parser.add_argument(
        "--output", type=Path, default=Path("examples") / "workbooks" / "phase2_sample.xlsx"
    )
    run_parser = subparsers.add_parser("run", help="Run supported direct_xlsx steps or explicit simulation")
    run_parser.add_argument("job_id")
    run_parser.add_argument(
        "--fake-services",
        action="store_true",
        help="Explicitly simulate steps; does not perform real Excel or email operations",
    )
    run_parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Process and validate in temporary storage without publishing or sending email",
    )
    return parser


def _load(path: Path) -> Configuration:
    return load_and_validate(path, plugin_ids=APPROVED_EXAMPLE_PLUGIN_IDS)


def _list_jobs(configuration: Configuration) -> int:
    print(f"Configuration valid: {configuration.source_path}")
    for job in configuration.jobs.values():
        reason = DirectXlsxServices.capability_error(job)
        capability = (
            f"production-executable=no ({reason})"
            if reason else "real-direct_xlsx=eligible (workbook preflight required)"
        )
        print(
            f"{job.id}: schema-valid; fake-executable; {capability}"
        )
    return EXIT_OK


def main(argv: list[str] | None = None) -> int:
    args = _parser().parse_args(argv)
    if args.command == "generate-samples":
        try:
            generate_sample(args.output.resolve())
        except (AutomationError, OSError) as exc:
            detail = str(exc) if isinstance(exc, AutomationError) else type(exc).__name__
            print(f"Sample generation failed: {detail}; no output overwritten", file=sys.stderr)
            return EXIT_RUNTIME
        print(f"Generated non-sensitive plain workbook: {args.output.resolve()}")
        return EXIT_OK
    try:
        configuration = _load(args.config)
    except ConfigurationError as exc:
        print(f"Configuration error: {exc}", file=sys.stderr)
        return EXIT_CONFIGURATION

    if args.command == "validate":
        print(f"Configuration valid: {configuration.source_path}")
        print(f"Validated {len(configuration.jobs)} job(s); none were executed.")
        return EXIT_OK

    if args.command == "list-jobs":
        return _list_jobs(configuration)

    if args.command == "run":
        job = configuration.jobs.get(args.job_id)
        if job is None:
            print(f"Unknown job ID: {args.job_id}", file=sys.stderr)
            return EXIT_CONFIGURATION
        services = FakeExecutionServices() if args.fake_services else DirectXlsxServices()
        try:
            logger = configure_logging()
            result = run_job(
                job,
                services,
                dry_run=args.dry_run,
                on_event=print,
                logger=logger,
            )
        except AutomationError as exc:
            print(
                f"Run failed: {exc}", file=sys.stderr,
            )
            logging.getLogger("excel_automation").error(
                "job=%s status=failed error_type=%s",
                job.id,
                type(exc).__name__,
            )
            return EXIT_RUNTIME
        except OSError as exc:
            logging.getLogger("excel_automation").error(
                "job=%s status=failed error_type=%s",
                job.id,
                type(exc).__name__,
            )
            print(f"Run failed: filesystem error ({type(exc).__name__})", file=sys.stderr)
            return EXIT_RUNTIME

        if result["fake_execution"]:
            print(
                "FAKE SIMULATION SUCCEEDED; no Excel workbook was read or modified, "
                "no real validation was performed, no email was sent, and no output "
                f"was published. run_id={result['run_id']}"
            )
        else:
            publication = (
                f"output published: {job.output}" if result["output_published"]
                else "dry-run: temporary output validated and removed; final output not published"
            )
            print(
                f"REAL DIRECT_XLSX RUN SUCCEEDED; {publication}; "
                f"no desktop Excel or email used. run_id={result['run_id']}"
            )
        return EXIT_OK

    return EXIT_CONFIGURATION


if __name__ == "__main__":
    raise SystemExit(main())
