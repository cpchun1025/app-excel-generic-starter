# Job Configuration

## Status and format

The implemented format is YAML with `schema_version: 1`. A
configuration file contains multiple named jobs under `jobs`. Parsing and
strict validation are implemented. Phase 2 adds limited real `direct_xlsx`
staging, plain-value API operations, validation, and publication.
Plugin, configured copy-mapping, macro, rendering, and email execution are
not implemented. The original complete fake examples are in
[`../config/jobs.example.yaml`](../config/jobs.example.yaml).
The real Phase 2 example is
[`../config/jobs.phase2.yaml`](../config/jobs.phase2.yaml); its tested workbook
profile and CLI procedure are in [DIRECT_XLSX.md](./DIRECT_XLSX.md).

Configuration contains no executable expressions and no secret values. The
framework validates the entire selected job before creating an output copy or
performing external side effects. Unknown fields should be rejected in version
1 so misspellings do not silently change behavior.

### Current execution classifications

- **Schema-valid:** strict parsing and validation passed. Input workbook
  existence and operational capability are not implied.
- **Fake-executable:** explicitly injected fake services can simulate the
  workflow shape for orchestration tests. Simulation never reads, modifies, or
  validates a real workbook.
- **Real-execution eligible:** the selected job uses `direct_xlsx`, only
  implemented `validate` steps, valid check shapes, and no enabled email.
  Real preflight must additionally pass for every configured input file.
  The CLI runs eligible jobs without `--fake-services`; incompatible jobs
  fail clearly before staging. This is not a promise that arbitrary production
  workbooks are supported.

The example plugin IDs are accepted through a small explicit validation
allowlist. It only validates references; it does not resolve or invoke Python
callables and is not the Phase 3 plugin system. `copy`, `python`, `macro`,
`validate`, and `email` are schema-valid step types. In Phase 2 only `validate`
has a real step implementation; the others are fake-only.

Commands are `python -m excel_automation.cli --config PATH validate`,
`list-jobs`, and `run JOB_ID --fake-services [--dry-run]`. The global
`--config` option precedes the subcommand. Omit `--fake-services` for real
eligible jobs. Incompatible jobs and runtime failures return `3`; invalid
configuration or unknown job returns `2`; successful validation, listing, fake
simulation, or real execution returns `0`. A fake run reports simulation
explicitly and never publishes. Real `--dry-run` stages/checks only in
temporary storage and never publishes. All modes reject existing final outputs.
`generate-samples [--output PATH]` generates a fake plain `.xlsx` without loading
job configuration and never overwrites an existing file.

## Top-level schema

| Field | Required | Meaning |
|---|---:|---|
| `schema_version` | Yes | Integer. Must be `1` for this schema. |
| `jobs` | Yes | Non-empty mapping of unique job IDs to job definitions. |

Job IDs and step IDs are stable identifiers matching a conservative pattern
such as `[a-z0-9][a-z0-9_-]*`. A job definition:

```yaml
inputs:                         # required, one or more named workbook paths
  source: path/to/source.xlsx
workbook:                       # required
  base: source                  # required; key in inputs
  output: output/result.xlsx    # required; path to final output
  adapter: direct_xlsx          # required: direct_xlsx | excel_desktop
steps:                          # required, non-empty ordered list
  - id: validate-result
    type: validate
    checks:
      - type: file_exists
email:                          # optional; omitted means no email activity
  enabled: false
```

### Job fields

| Field | Required | Meaning / constraints |
|---|---:|---|
| `inputs` | Yes | Mapping of aliases to workbook paths. Paths are resolved as described below. One input may also serve as `workbook.base`. |
| `workbook.base` | Yes | Must refer to a key in `inputs`. Real execution stages a temporary copy and publishes to `workbook.output` only after success and only outside dry-run. |
| `workbook.output` | Yes | Output path. Must not resolve to any input path. Existing-output policy is fail-on-collision before any run, including fake and dry runs. |
| `workbook.adapter` | Yes | `direct_xlsx` or `excel_desktop`; compatibility is checked before staging. |
| `steps` | Yes | Non-empty list, executed in listed order. Each step has unique `id` and supported `type`. |
| `email` | No | If omitted or `enabled: false`, no message is prepared or sent. If enabled, exactly one final email step is required. |

## Ordered steps

Steps have a unique `id`, a `type`, and only fields valid for that type.
References to input aliases and output workbook are checked before execution.

### `copy`

Copies a source sheet/range into a destination sheet/range using the job's
selected adapter. Intended to be placed before a macro if that macro consumes
the copied data.

| Field | Required | Meaning |
|---|---:|---|
| `id` | Yes | Unique step ID. |
| `type` | Yes | `copy`. |
| `source.workbook` | Yes | Input alias. |
| `source.sheet` | Yes | Exact source sheet name. |
| `source.range` | Yes | A1-style range, e.g. `A1:C20`. |
| `destination.workbook` | Yes | `output`. |
| `destination.sheet` | Yes | Exact destination sheet name. |
| `destination.range` | Yes | A1-style top-left destination range or equal-size target, as defined by the implementation. |
| `mode` | Yes | `values`, `values_and_formatting`, or `worksheet`. Adapter must explicitly support requested mode or reject before mutation. |

For version 1, a range copy must validate dimensions and destination bounds.
Whole-worksheet copy has no range fields. To avoid ambiguity, `mode:
worksheet` requires `source.range` and `destination.range` to be omitted and
requires the adapter to support whole-sheet copy. Do not mix range and
worksheet fields.

### `python`

Calls a registered Python plugin.

| Field | Required | Meaning |
|---|---:|---|
| `id` | Yes | Unique step ID. |
| `type` | Yes | `python`. |
| `plugin` | Yes | Approved plugin ID; never a Python import path. |
| `options` | No | Plugin-specific mapping; defaults to `{}`. Plugin validates its own options. |

Plugin call contract and registration are defined in
[`ARCHITECTURE.md`](./ARCHITECTURE.md). The runner rejects unknown plugin IDs
before output staging.

### `macro`

Runs a named VBA macro in the staged output workbook.

| Field | Required | Meaning |
|---|---:|---|
| `id` | Yes | Unique step ID. |
| `type` | Yes | `macro`. |
| `workbook` | Yes | Must be `output`; macro runs against the staged output. |
| `name` | Yes | Explicit fully qualified macro name, e.g. `Module1.Refresh`. |
| `args` | No | Ordered list of scalar arguments; defaults to `[]`. |

`macro` is incompatible with `adapter: direct_xlsx`. It requires Windows and
installed desktop Excel. It does not mean that an `.xlsx` file has a macro.
Only trusted macros may run, and macro security must not be globally weakened.

### `validate`

Runs one or more checks. Any failed check fails the job and blocks all later
email work.

| Field | Required | Meaning |
|---|---:|---|
| `id` | Yes | Unique step ID. |
| `type` | Yes | `validate`. |
| `checks` | Yes | Non-empty list of checks. |

Proposed check types:

| Check | Required fields | Meaning |
|---|---|---|
| `file_exists` | none | Staged output file exists and is non-empty. |
| `sheet_exists` | `sheet` | Named sheet exists in output. |
| `range_nonempty` | `sheet`, `range` | Configured range contains at least one non-empty cell. |
| `range_equals` | `sheet`, `range`, `expected` | Range value equals expected scalar or matrix; useful for fake fixtures. |

Phase 2 implements these four checks for the constrained `direct_xlsx`
profile. `file_exists` inspects the staged output before publication, not the
not-yet-published final path. Fake services still simulate check outcomes.
`range_equals` accepts a scalar for one cell or a rectangular two-dimensional
matrix matching the selected range; a flat list or mismatched matrix fails
real capability preflight. Numeric equality allows integral floats to equal
integers but booleans do not equal numbers. A1 ranges must be forward and
within Excel bounds and the adapter limits in [DIRECT_XLSX.md](./DIRECT_XLSX.md).
Version 1 schema validation requires validation before any `email` step;
real email remains unimplemented.

### `email`

Prepares a MIME message and, if enabled and not dry-run, may send it.
Configuration is top-level `email` as well as a final ordered step:

```yaml
email:
  enabled: true
  send: false
  smtp:
    host: smtp.invalid.example
    port: 587
    tls: starttls
    sender: automation@example.invalid
    recipients:
      - report-reader@example.invalid
    username: ${EXCEL_SMTP_USER}
    password: ${EXCEL_SMTP_PASSWORD}
  subject: "Fake report: {job_id}"
  body: "Run {run_id} completed."
  delivery_key: "{job_id}:{scheduled_occurrence}"
  outputs:
    - type: html_table
      sheet: Summary
      range: A1:C8
      placement: body
    - type: workbook
      placement: attachment
```

The final `email` step contains only `id` and `type: email`; delivery details
live in the top-level section to avoid duplicating SMTP settings.

| Field | Required when `enabled: true` | Meaning |
|---|---:|---|
| `enabled` | Yes | Boolean. False means no email preparation or sending. |
| `send` | No | Boolean; defaults to `true`. When false, prepare the message only and do not resolve credentials or connect to SMTP. `--dry-run` always behaves as if this were false. |
| `smtp.host` | Yes | SMTP host. Fake examples use reserved `.invalid` names. |
| `smtp.port` | Yes | Integer port, valid range 1–65535. |
| `smtp.tls` | Yes | `none`, `starttls`, or `implicit`. |
| `smtp.sender` | Yes | Sender address. |
| `smtp.recipients` | Yes | Non-empty list of recipient addresses. |
| `smtp.username` | No | Literal-free secret reference, usually `${ENV_VAR}`. |
| `smtp.password` | No | Literal-free secret reference, usually `${ENV_VAR}`. |
| `subject` | Yes | Template supporting documented safe metadata placeholders only. |
| `body` | Yes | Text template; spreadsheet cells are not interpolated unsafely. |
| `delivery_key` | Yes | Stable idempotency key template unique to logical delivery. |
| `outputs` | Yes | Non-empty list of email output modes. |
| `outputs[].type` | Yes | `html_table`, `range_image`, `html_attachment`, or `workbook`. |
| `outputs[].placement` | Yes | `body` for HTML table, `attachment` for image/file/workbook. |
| `outputs[].sheet`, `range` | For range-based modes | Exact sheet/range in validated output workbook. |

`html_table` is body content. `range_image` and `html_attachment` must be
attachments. `workbook` attaches the completed output and must be an
attachment. Each output mode's adapter compatibility is in the matrix below.
TLS modes are mutually exclusive. Username/password must be absent together or
provided together. Credential values must never appear literally in YAML,
logs, or committed files.

An enabled email requires exactly one `email` step, after all processing and
validation steps, and it must be the last step. An email step when email is
disabled is invalid. Validation cannot follow email. `send: false` is intended
to prepare email content without resolving secrets or contacting SMTP. In
Phases 1 and 2 neither email preparation nor transmission is implemented; fake
services only record a simulated email step. `--dry-run` always suppresses
SMTP regardless of configuration; email content may be prepared and
summarized in a future implementation but not transmitted.

## Cross-field and ordering constraints

- A job has exactly one `workbook.base`, one output path, and one adapter.
- Base input and output paths must differ after normalization.
- `macro` requires `excel_desktop`; `direct_xlsx` must reject `.xlsm` or any
  workbook feature it cannot preserve.
- `worksheet` copy mode cannot include ranges; other copy modes require source
  and destination ranges.
- A `python` step's plugin ID must exist in the trusted registry.
- A `validate` step must occur after mutations it validates.
- If email is enabled, at least one validation step precedes the final email
  step. Processing or validation failure prevents email preparation and send.
- Email output ranges must refer to the staged output and use only supported
  renderers.
- Unknown fields, unsupported step types, invalid references, duplicate step
  IDs, path collisions, and incompatible adapter/feature combinations fail
  preflight.

## Path and secret resolution

Relative configuration and workbook paths resolve against the directory
containing the YAML file, not the process current working directory. Absolute
paths remain absolute after normalization. Output directories are created
only after real preflight succeeds and only for normal publication. Dry runs
and fake runs do not create output directories; inputs are never silently
created. Normal publication uses a complete staged file and atomic
no-overwrite hard-link creation. It requires a filesystem supporting hard
links; there is no partial-copy fallback.
The application/repository root for `logs/` is resolved independently from
the current working directory.

Secrets are referenced with `${ENV_VAR}` tokens or another explicitly
implemented secret-provider URI. Version 1 proposes environment-variable
references. Resolve secrets only immediately before SMTP connection; fail
clearly when missing. Never interpolate secret values into logs, error
messages, reports, or example configuration. Do not commit real credentials.

## Feature / adapter compatibility and implementation status

This table distinguishes the configuration contract from implementation
status. A constrained real `.xlsx` adapter exists; no desktop Excel,
renderer, callable plugin, or SMTP adapter exists.

| Feature | `direct_xlsx` target | `excel_desktop` target | Status now |
|---|---|---|---|
| Config parsing/preflight | Adapter-independent | Adapter-independent | Implemented in Phase 1 |
| Fake-adapter orchestration | Test-only fake | Test-only fake | Implemented in Phase 1; simulation only |
| Simple `.xlsx` values | Only the strict profile in DIRECT_XLSX.md | Not implemented | Phase 2: staged-copy read/write API tested; no configurable write step |
| Values copy | Supported `.xlsx` only | Excel-supported workbooks | Not implemented |
| Copy with formatting | No claim until verified | Candidate; integration verification required | Not implemented |
| Whole-worksheet copy | No claim until verified | Candidate; integration verification required | Not implemented |
| Python plugin-ID validation | Independent of adapter | Independent of adapter | Implemented as explicit example/test allowlist only |
| Python plugin execution | Independent of adapter; plugin services are injected | Independent of adapter; plugin services are injected | Not implemented |
| VBA macro execution | Unsupported | Windows + installed desktop Excel required | Not implemented |
| File/sheet/range validation | Four checks against staged plain `.xlsx` output | Not implemented | Implemented in Phase 2 |
| Real staging/publication | Byte copy; atomic no-overwrite final name | Not implemented | Implemented in Phase 2; dry-run never publishes |
| HTML table / HTML attachment | From supported values, escaped | May use Excel-derived values if needed | Not implemented |
| Range image | Unsupported | Windows + installed desktop Excel in proposed approach | Not implemented |
| MIME construction | Independent of Excel | Independent of Excel | Not implemented |
| SMTP send | Independent of Excel | Independent of Excel | Not implemented |
| SMTP credentials | Secret provider, not adapter-bound | Secret provider, not adapter-bound | Not implemented |
| Idempotent email ledger | Persistent storage required | Persistent storage required | Not implemented |

See [`IMPLEMENTATION_PLAN.md`](./IMPLEMENTATION_PLAN.md) for acceptance gates.
