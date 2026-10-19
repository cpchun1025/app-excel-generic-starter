# Implementation Plan

## Delivery approach

Implement a minimal reusable framework in independently testable increments.
Each phase adds a capability only after the previous phase's acceptance
criteria pass. Example configurations are illustrative until their matching
phase is complete. Do not invent production workbook rules or claim support
for unimplemented operations.

## Phase 1 — Configuration, CLI, logging, orchestration, fake adapters

**Historical boundary decisions (Phase 1; Phase 2 changes below)**

- Schema-valid means YAML and cross-field constraints validate; it does not
  imply file existence or a real supported operation.
- Fake-executable means the job shape can run through explicitly injected
  fake services. Production-executable means real adapters perform the work;
  no jobs are production-executable in Phase 1.
- Example jobs validate, but configured copy, Python, macro, validation, and
  email work is simulated only. No real operation is implemented.
- Fake execution always uses a temporary workspace and never publishes
  `workbook.output`, including normal fake runs. `--dry-run` also uses only
  temporary staging and never publishes.
- Output collision checks run before any fake or dry run: if the configured
  final output already exists, the run fails. Output/input path identity is
  rejected during configuration validation.
- Future production `file_exists` checks the staged file before final
  publication. Phase 1 fake validation simulates this against a virtual
  staged-output token.
- Phase 1 accepts the two example plugin IDs using a small explicit allowlist.
  This validates references only; no callable resolution, plugin contract, or
  real Python transformation is implemented until Phase 3.

**Scope**

- Establish package, dependency declarations, typed configuration models,
  YAML loading, and fail-fast validation for version 1.
- Implement `list-jobs` and `run JOB` command shapes, including `--dry-run`,
  concise summaries, and documented exit codes.
- Configure console and rotating repository-root `logs/` logging with run ID,
  job, step status, safe exception context, and secret/cell-content redaction.
- Implement orchestration lifecycle and ordered step dispatch with injected
  fake services. Fake services only record/simulate work; they do not read or
  write workbooks, run plugins/macros, validate actual content, or send email.
- Add fake output staging and failure/cleanup semantics.

**Acceptance tests**

- Valid Python-only, macro-only, and mixed fake configurations parse.
- Invalid schema version, unknown fields/jobs/steps, duplicate IDs, missing
  references, invalid ordering, plugin IDs not registered, incompatible
  adapters, and path collisions fail before output staging.
- Fake steps execute exactly in configured order; recorded logs include run,
  job, and step outcomes but omit configured secret tokens and fake cell
  contents.
- Injected step failure prevents later validation/email; validation failure
  prevents email preparation and send.
- Dry-run never calls the fake transport's send method.
- CLI returns documented success/configuration/runtime failure exit codes.

**Done when**

Unit tests cover all acceptance items and pass; `list-jobs`, a fake `run`, and
fake `--dry-run` are demonstrable without Excel or network access. The Phase 1
test command is `python -m pytest -q`.

## Phase 2 — Direct `.xlsx` workbook operations

**Implemented boundary**

- Status: Phase 2 implemented and portable acceptance tests verified with
  `python -m pytest -q -W error` (78 passed, including Phase 1 regression tests).
- Version 1 already expresses a real staging-and-validation job: staging is
  implicit and `validate` is an existing step. No schema or version change
  is needed; see [jobs.phase2.yaml](../config/jobs.phase2.yaml).
- The real runtime accepts only `direct_xlsx` with implemented validation
  checks. It rejects the entire selected job before staging if any step is
  a plugin, configured copy mapping, macro, rendering, or email operation.
  Schema-valid does not mean real-executable.
- Adapter-level plain-value read/write is implemented and tested separately;
  no configurable write or plugin step has been added.
- Normal real success publishes the validated staged file only. Fake execution
  remains temporary-only. Real `--dry-run` runs real checks against a temporary
  copy, removes it, never publishes, and has no email service.
- Existing final outputs fail in every mode; publication uses an atomic
  no-overwrite hard link on the same filesystem. Filesystems without hard-link
  support fail explicitly instead of falling back to a partial final copy.
- The supported workbook profile, limits, rejected features, tests, and real
  CLI example are documented in [DIRECT_XLSX.md](./DIRECT_XLSX.md).

**Scope**

- Implement direct adapter support only for explicitly scoped `.xlsx`
  value reads/writes, safe whole-file staging copies, file/sheet/range checks,
  and sample workbook generation.
- Reject `.xlsm`, formatting/worksheet features, or unsupported workbook
  features that the adapter cannot preserve.
- Do not implement or imply VBA execution in this phase.

**Acceptance tests**

- Generate small fake `.xlsx` files and test normal, empty, malformed,
  duplicate, and boundary cases where applicable.
- Verify source/template files remain byte-for-byte unchanged by job runs.
- Verify whole-file staging copy, adapter value read/write, output file
  creation, sheet existence, and range checks. Configured values range
  copying remains Phase 4 and is rejected in Phase 2.
- Unsupported macro-enabled or feature-rich inputs are rejected explicitly,
  not rewritten or silently degraded.

**Done when**

Tests pass for documented `.xlsx` capabilities and unsupported inputs fail
clearly. Documentation lists tested formats and limitations. Portable tests
run with `python -m pytest -q`; no desktop Excel or SMTP is used.

## Phase 3 — Python plugin API and sample plugins

**Scope**

- Implement typed plugin context/result contract and explicit trusted
  plugin-ID registry.
- Add one or two fake sample plugins and keep transformations separable from
  workbook I/O.
- Validate plugin IDs/options before staging output where possible.

**Acceptance tests**

- Test plugin logic with unit-level values and generated fake workbooks.
- Verify registered IDs execute and unknown IDs fail preflight.
- Verify config cannot request arbitrary imports or code execution.
- Verify plugin exceptions fail only the current run and prevent later steps.

**Done when**

Plugin API is documented and tested, and the Python-only example runs end to
end using only supported direct `.xlsx` features.

## Phase 4 — Copy mappings and output validation

**Scope**

- Implement configured source/destination sheet and range mappings for
  supported direct-adapter `values` copies.
- Implement supported validation checks and strict check ordering.
- Keep `values_and_formatting` and `worksheet` modes rejected until their
  specific adapter implementation and tests land.

**Acceptance tests**

- Test matching/mismatched dimensions, empty ranges, missing sheets,
  invalid A1 references, output range checks, and protected source files.
- Test all copy and validation failure paths prevent downstream email work.
- Test error output does not expose copied spreadsheet cell contents.

**Done when**

Documented copy and validation modes work in pytest fixtures and every
unsupported mode fails before modifying output.

## Phase 5 — Windows desktop Excel and macro automation

**Scope**

- Implement Windows-only desktop Excel adapter using explicit workbook paths.
- Support staged `.xlsm` output, exact macro names and arguments, resource
  cleanup on success/failure, and clear diagnostics.
- Never weaken macro security globally or run untrusted macros.
- Keep integration tests opt-in and separate from portable unit tests.

**Acceptance tests**

- Unit tests mock the Excel automation boundary, including open, target
  workbook selection, macro invocation, errors, close, and release.
- Separate integration procedure requires a user-supplied trusted `.xlsm`
  with an agreed test macro and verifies actual output.
- Verify source/template is unchanged; target macro runs on configured output
  even if another workbook is active.
- Test missing workbook/sheet/macro, locked file, Excel failure, and cleanup.

**Done when**

Mocked tests pass on ordinary development platforms; documented real-Excel
integration passes on Windows with desktop Excel and a user-supplied fixture.
No generated `.xlsx` is represented as macro-enabled.

## Phase 6 — Rendering and email message preparation

**Scope**

- Add escaped HTML table generation and HTML-file attachment without Excel
  where supported.
- Add MIME message construction and workbook attachment.
- Add range-image rendering through desktop Excel only after documenting its
  environment and rendering limitations.
- Keep preparation separate from SMTP transmission.

**Acceptance tests**

- Test HTML escaping, empty ranges, MIME types, output modes, attachments,
  malformed ranges, and no accidental inclusion of secrets.
- Mock Excel image rendering; separately verify image output in Windows/Excel
  integration tests.
- Verify a failed processing or validation step results in no prepared email.

**Done when**

Message construction is fully unit-tested without SMTP or Excel, and each
Excel-dependent rendering capability has a separate integration test.

## Phase 7 — SMTP, secrets, idempotency, and operational documentation

**Scope**

- Implement approved secret-reference resolution, SMTP TLS modes, send
  transport, and persistent idempotency ledger.
- Enforce `--dry-run` at the SMTP boundary.
- Implement at-most-once automatic behavior: sent keys are skipped; uncertain
  send outcomes are marked `unknown` for operator reconciliation.
- Document ledger location, retention, single-host limits, deployment and
  scheduling considerations.

**Acceptance tests**

- Mock SMTP for successful send, rejection, timeout/uncertain outcome, TLS
  setup, missing secrets, and redaction.
- Verify dry-run cannot invoke SMTP even if configuration enables sending.
- Verify duplicate delivery keys cannot produce a second send after recorded
  success and uncertain outcomes do not auto-retry.
- Verify no send occurs after processing/rendering/validation errors.

**Done when**

SMTP tests use only mocks, idempotency tests pass against the selected
persistent store, and no integration/unit test sends real email.

## Phase 8 — End-to-end examples and release readiness

**Scope**

- Generate fake `.xlsx` sample workbooks, finalize example jobs and README,
  add installation/troubleshooting/security notes, CLI help, sample generation,
  test commands, and separate Excel integration instructions.
- Ensure each example uses only capabilities implemented by that release.
- Keep fake `.xlsm` references clearly marked as placeholders for
  user-supplied trusted integration fixtures.

**Acceptance tests**

- Run full portable pytest suite from a clean environment.
- Run CLI list, one Python-only example, and dry-run without Excel or network.
- Confirm examples fail with clear prerequisites if a Windows/Excel feature is
  invoked on an unsupported platform.
- Audit tracked configuration and logs for credentials or sensitive values.

**Done when**

All required deliverables are documented, portable tests pass, Excel tests
remain separately opt-in, examples accurately describe their implementation
status, and no test sends a real email.
