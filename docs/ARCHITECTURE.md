# Framework Architecture

## Scope and status

This document describes the reusable framework design and the currently
implemented Phase 1 and Phase 2 subsets. Capability tables in this document and
[CONFIGURATION.md](./CONFIGURATION.md) distinguish design targets from
implemented behavior.

The framework runs named jobs from versioned configuration. Configuration
selects approved operations and plugins; it is data, not executable Python.
Job-specific business rules live in tested, trusted plugins.

## Proposed project tree

```text
.
├── config/
│   ├── jobs.example.yaml
│   └── jobs.phase2.yaml
├── docs/
│   ├── ARCHITECTURE.md
│   ├── CONFIGURATION.md
│   ├── IMPLEMENTATION_PLAN.md
│   ├── DIRECT_XLSX.md
│   └── integration-testing.md
├── examples/
│   ├── plugins/
│   │   └── sample_jobs.py
│   └── workbooks/
│       └── (generated fake .xlsx fixtures; no fake .xlsm macros)
├── logs/
│   └── (runtime rotating logs; ignored by version control)
├── src/
│   └── excel_automation/
│       ├── __init__.py
│       ├── cli.py
│       ├── errors.py
│       ├── samples.py
│       ├── config/
│       │   ├── loader.py
│       │   ├── models.py
│       │   └── validation.py
│       ├── runner/
│       │   ├── orchestrator.py
│       │   ├── context.py
│       │   ├── direct_services.py
│       │   ├── steps.py
│       │   └── services.py
│       ├── workbook/
│       │   ├── protocol.py
│       │   ├── direct_xlsx.py
│       │   ├── excel_desktop.py
│       │   ├── copies.py
│       │   └── validation.py
│       ├── plugins/
│       │   ├── example_allowlist.py
│       │   ├── api.py
│       │   └── registry.py
│       ├── macros/
│       │   └── executor.py
│       ├── rendering/
│       │   ├── html.py
│       │   └── excel_image.py
│       ├── email/
│       │   ├── message.py
│       │   ├── idempotency.py
│       │   └── smtp.py
│       └── logging_config.py
├── tests/
│   ├── unit/
│   └── integration/
│       └── test_excel_desktop.py
├── pyproject.toml
└── README.md
```

**Phase 1 present now:** `cli.py`, `errors.py`, the three `config` modules,
`logging_config.py`, and `runner/context.py`, `runner/orchestrator.py`, and
`runner/services.py`, plus `plugins/example_allowlist.py` for ID validation
only. Phase 2 additions are listed next; remaining module entries in the tree
are proposed future modules.

**Phase 2 present now:** `workbook/direct_xlsx.py`, `workbook/copies.py`,
`workbook/validation.py`, `runner/direct_services.py`, and `samples.py`.
Configured copy mappings, callable plugins, desktop Excel, rendering, and
SMTP remain future work. `workbook/copies.py` currently stages/publishes
whole files only, not sheet/range mappings. See [DIRECT_XLSX.md](./DIRECT_XLSX.md)
for the exact accepted package profile and rejection rules.

### Module responsibilities

| Module | Responsibility |
|---|---|
| `cli.py` | `list-jobs`, `validate`, `generate-samples`, and `run` (real constrained direct_xlsx or explicit fake); concise progress and stable exit codes. |
| `config/loader.py` | Read YAML and identify source/configuration errors. |
| `config/models.py` | Typed versioned models for jobs, inputs, outputs, steps, checks, and email. |
| `config/validation.py` | Validate schema, references, step ordering constraints, incompatible options, and registered plugin IDs before side effects. |
| `runner/orchestrator.py` | Create run context, stage output, execute ordered steps, gate validation/email, and produce a run result. |
| `runner/context.py` | Immutable run/job metadata and injected services supplied to steps/plugins. |
| `runner/steps.py` (future; `runner/services.py` currently injects fake services) | Dispatch typed step models to the matching service without executing arbitrary configuration code. |
| `workbook/protocol.py` | Adapter interface for workbook inspection, supported edits, copy mappings, save, and validation. |
| `workbook/direct_xlsx.py` | Direct `.xlsx` operations limited to explicitly supported features; never executes VBA. Must reject unsupported feature-preservation requirements rather than silently discard them. |
| `workbook/excel_desktop.py` | Windows desktop Excel automation for Excel-dependent operations, with explicit workbook paths and deterministic close/release. |
| `workbook/copies.py` | Implemented: whole-file staging and atomic no-overwrite publication. Configured sheet/range copies are future work. |
| `runner/direct_services.py` | Preflight every real step and input; reject unsupported operations before staging; execute real Phase 2 checks. |
| `samples.py` | Generate non-sensitive plain `.xlsx` fixtures without overwriting files. |
| `workbook/validation.py` | Run configured file, sheet, range, and other implemented output checks. |
| `plugins/api.py` | Define `JobContext`, plugin callable, and typed `PluginResult` contract. |
| `plugins/registry.py` | Register and resolve reviewed plugin IDs to trusted Python callables. |
| `plugins/example_allowlist.py` | Phase 1 explicit approved-ID set for schema checks only; not callable registration. |
| `macros/executor.py` | Validate the intended macro target and delegate execution to the Excel adapter; no macro execution via `openpyxl`. |
| `rendering/html.py` | Escape cell text and build supported HTML tables/files without Excel where possible. |
| `rendering/excel_image.py` | Render selected ranges through desktop Excel; limitations and prerequisites must be documented. |
| `email/message.py` | Construct MIME content and attachments without sending. |
| `email/idempotency.py` | Track delivery keys and prevent duplicate sends across retry/restart. |
| `email/smtp.py` | Resolve approved secret references and send MIME messages via SMTP. |
| `logging_config.py` | Configure console and rotating repository-root `logs/`; application messages use safe metadata only. |

## Plugin contract and registration

Plugins use a small synchronous contract. An initial contract proposal is:

```python
from collections.abc import Mapping
from typing import Protocol


class JobPlugin(Protocol):
    def __call__(
        self, context: "JobContext", options: Mapping[str, object]
    ) -> "PluginResult": ...
```

`JobContext` exposes a run ID, job ID, immutable resolved input paths, the
staged output path, the injected workbook service, and a safe logger. It does
not expose SMTP credentials. `PluginResult` carries explicit success metadata
such as non-sensitive counts or named outputs; failures are raised as typed
plugin exceptions and fail the step. Cell values and secrets must never be
copied into logs or exception summaries.

Configuration refers only to an approved stable ID, for example
`sample.clean_sales`. The framework resolves that ID through an explicit
registry, such as:

```python
PLUGIN_REGISTRY = {
    "sample.clean_sales": sample_jobs.clean_sales,
    "sample.prepare_forecast": sample_jobs.prepare_forecast,
}
```

The registry is application code reviewed and shipped with the deployment (or
an explicitly configured trusted extension registry). It is not assembled by
importing module paths from a job file, evaluating expressions, or executing
code embedded in configuration. Unknown IDs fail configuration validation
before any workbook is copied or modified. In Phase 1, validation uses only a
small explicit allowlist of example/test IDs (`sample.clean_sales`,
`sample.prepare_forecast`). This confirms the reference is approved, but does
not resolve or invoke a callable and is not implementation of the Phase 3
plugin API. Fake services only record that the configured plugin step was
simulated.

## Job execution lifecycle

1. **Load and select:** parse configuration, select one named job, and reject
   unknown schema versions or job IDs.
2. **Preflight validation:** validate all fields, step ordering, input/output
   references, adapter compatibility, plugin IDs/options, paths, and required
   secret *references*. Do not read or log secret values at this stage.
3. **Initialize run:** create a unique run ID; configure console and rotating
   logs under repository `logs/`; record job and run metadata without workbook
   contents.
4. **Stage output:** production design copies the configured base workbook to a
   unique staging location. In Phase 1, fake services create only a temporary
   workspace and virtual staged-output token; they do not read or copy
   workbook bytes. Inputs and templates remain unchanged.
5. **Execute ordered steps:** execute each validated step in configuration
   order, marking start/success/failure. `copy`, `python`, and `macro` steps
   use their selected services. Macro steps use the Excel adapter only.
6. **Validate output:** run required configured checks. Validation failure
   fails the job and prevents email preparation and delivery.
7. **Prepare email:** render requested body/attachments and construct a MIME
   message only after processing and validation succeeded. `--dry-run` stops
   after preparation/summary and categorically does not open an SMTP
   connection.
8. **Send (if enabled):** derive a stable delivery key for the logical job
   occurrence and atomically claim it in a persistent delivery ledger before
   sending. A repeated key already marked sent is skipped. Record outcome and
   report ambiguous outcomes rather than blindly resending.
9. **Finalize:** close/release adapter resources on success and failure,
   preserve diagnostic logs, report output path and result, and return the
   documented exit code.

### Phase 1 execution classes (historical)

The fake-only restrictions below describe Phase 1. The following Phase 2
transition section supersedes them for eligible real `direct_xlsx` jobs.

- **Schema-valid:** configuration passes parsing and strict schema/cross-field
  validation. This says nothing about file existence, Excel, or actual
  operation support.
- **Fake-executable:** the ordered job can be passed through explicitly
  injected fake services for orchestration testing. A fake success is always
  labeled a simulation and never represents workbook processing.
- **Production-executable:** requires real adapters and operations. No job is
  production-executable in Phase 1, including jobs whose YAML validates.

`list-jobs` reports the first two statuses and explicitly reports production
execution unavailable. `run` refuses unless `--fake-services` is supplied.
The fake bundle does not read workbook inputs, edit workbooks, execute plugin
functions or VBA, perform real validations, render content, resolve secrets,
or send email.

Phase 1 CLI examples (run from the repository root):

```text
python -m pip install -e ".[dev]"
python -m excel_automation.cli --config config/jobs.example.yaml validate
python -m excel_automation.cli --config config/jobs.example.yaml list-jobs
python -m excel_automation.cli --config config/jobs.example.yaml run fake-python-only --fake-services
python -m excel_automation.cli --config config/jobs.example.yaml run fake-python-only --fake-services --dry-run
python -m pytest -q
```

Exit codes are `0` for successful validation/list/simulation, `2` for invalid
configuration or an unknown job, and `3` when production execution is
requested or a fake run fails.

### Phase 2 transition

The preceding fake-only restriction describes Phase 1. Phase 2 permits real
execution without `--fake-services` only if every selected step is implemented
and compatible with `direct_xlsx`. Currently that means staging plus
`validate` steps. This requires no schema change. Real input inspection runs
before staging; unsupported steps or workbook features fail rather than being
skipped. The three original plugin/macro examples remain fake-only.

Normal real execution stages on the output filesystem, checks the staged
workbook, and publishes only after success using atomic no-overwrite
hard-link creation. Real dry-run stages in temporary storage, performs real
checks, and never publishes or sends email. Both modes reject an existing
final output. Failure before publication leaves no completed final output.
The runner never edits source/template files.

The runtime bundle now explicitly declares `fake_execution` and a `preflight`
operation; it is injected into the same runner. Fake and real summaries are
distinct. `list-jobs` reports real capability eligibility, not successful
workbook preflight. See [DIRECT_XLSX.md](./DIRECT_XLSX.md) for CLI examples,
file-format limits, publication failure behavior, and test coverage.

### Failure behavior

Any preflight, staging, processing, macro, rendering, or required validation
failure stops later steps. In particular, no email is prepared/sent when a
required processing or validation step fails. Failure does not trigger blind
retries of a macro or email. Retry only explicitly classified safe operations.
Temporary artifacts are cleaned when safe; failed outputs/logs may be retained
for diagnosis according to a documented retention policy. Cleanup errors are
reported and do not mask the original failure.

Excel automation must always open the configured output workbook and invoke
the configured macro on it, never rely on the active workbook/window. Excel
resources must close in `finally`-equivalent cleanup, including failures.
Macro security must not be globally weakened; only trusted, approved workbooks
and macros may run.

## Adapters and operating-system boundary

The adapter choice is explicit or derived only from validated workbook
requirements; it is never silently changed to an implementation that may
damage workbook features.

| Operation | Direct `.xlsx` adapter | Windows desktop Excel adapter |
|---|---|---|
| Read/write simple `.xlsx` cell values | Intended for supported `.xlsx` cases; implementation and feature limits must be tested. | Available where Excel behavior/features are required. |
| Run VBA | Not supported. No Python workbook library is a VBA execution engine. | Requires Windows and installed desktop Microsoft Excel; opens the explicit `.xlsm` output and invokes the exact macro. |
| Preserve VBA project in `.xlsm` | Not supported by the direct `.xlsx` adapter. Reject rather than risk stripping VBA. | Intended for macro-enabled workbooks; integration-test actual preservation and execution. |
| Copy values | Intended for supported `.xlsx` ranges. | Available for Excel-dependent files/workbooks. |
| Copy values and formatting / whole worksheet | Not claimed until exact feature coverage is implemented and tested; reject unsupported requests. | Candidate for Excel automation; exact semantics and workbook-feature fidelity require integration testing. |
| Validate file/sheet/range | File checks can be library-independent; workbook structure checks depend on adapter and supported formats. | Available through Excel for compatible workbooks. |
| HTML table / HTML attachment | Can be generated from supported cell values without Excel; escape cell content. | Can extract values where workbook calculation/render behavior requires Excel. |
| Range image rendering | Not supported by direct `.xlsx` adapter. | Requires Windows desktop Excel in the proposed approach; rendering can vary with Excel version, fonts, display configuration, and page setup. |
| SMTP | Not an adapter responsibility and does not require Excel. | Not an adapter responsibility and does not require Excel. |

This table describes intended adapter boundaries, not blanket implementation
claims. Plain `.xlsx` values and staged checks are now implemented only for
the strict Phase 2 profile; no other Excel features are implied.
See the implementation status boundary in [CONFIGURATION.md](./CONFIGURATION.md).

## Output copies, validation, and dry-run

The input/template is immutable by default. Production execution will create
a distinct output copy before mutation and direct every mutating step to that
copy. Phase 1 has no real staging or publication; Phase 2 implements the
constrained real transition described above.

Phase 1 fake execution uses an isolated temporary workspace for every run,
whether or not `--dry-run` is present; it never creates `workbook.output`.
`--dry-run` additionally marks the simulated run as dry-run but does not
change publication behavior: nothing is published in either fake mode. Before
starting either mode, the runner rejects an output path that already exists.
The collision check is against the configured final destination even though
the fake run only uses temporary storage. Configured output must also differ
from all input paths after normalization.

In Phase 2 real execution, `file_exists` checks the staged output while
the job is running, before final publication, not the not-yet-published final
path. In Phase 1, fake validation services simulate the check against a
virtual staged-output token and do not inspect the filesystem. The validation
result must therefore be understood only as a fake orchestration outcome.

Configured validation is a gate: required output file, expected sheet/range,
or other implemented checks must pass before email preparation. A successful
Phase 1 dry run means only that injected fake services completed their steps.
It does not mean processing, real validation, or email preparation occurred.
It never means an email was sent. Future production `--dry-run` must be
enforced at the transport boundary as well as the CLI.

## Logging and duplicate-email prevention

Use a rotating log file under `logs/` resolved from the installed repository/
application root, not the current working directory, plus concise console
output. Include timestamp, job name, run ID, step, and outcome. Redact secret
values and never log spreadsheet cell contents, credentials, full email
bodies, or attachment contents. Errors should include useful exception type
and safe context.

SMTP delivery has an unavoidable ambiguity if a connection drops after the
server accepted a message but before the client recorded success. The proposed
policy is **at-most-once automatic sending** using a persistent local delivery
ledger: create a deterministic key from job ID, configured business run key
(or explicit scheduled occurrence), and message identity; claim it atomically
before send; record `sent` only after SMTP success. On uncertain outcome,
record `unknown` and require operator reconciliation instead of retrying
automatically. A ledger prevents framework retries from resending a known
successful send, but cannot provide universal exactly-once delivery across an
SMTP server and local storage. Ledger storage, retention, and multi-host
coordination must be selected before multi-instance deployment.
