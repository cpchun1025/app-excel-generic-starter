# Phase 2: constrained real `.xlsx` execution

## Implemented versus specified

Phase 2 adds real file staging, plain-value adapter read/write methods,
file/sheet/range checks, generated sample workbooks, and no-overwrite output
publication. No desktop Excel installation is needed; portable Python and the
declared dependencies are sufficient.

The existing schema can express a real job without a new step type: the runner
stages a copy automatically, then executes one or more `validate` steps.
[jobs.phase2.yaml](../config/jobs.phase2.yaml) demonstrates that job. This
example really copies and checks a workbook; it does not transform data.
Adapter value read/write is exercised directly in tests, not exposed as a new
configuration step.

Plugin execution, configured copy mappings (even `mode: values`), macros,
rendering, and email remain unimplemented. Schema validation can accept
references to them, but real execution rejects the whole job before creating
a workspace or staging a file. Fake mode remains explicitly opt-in and never
claims a real workbook result. `list-jobs` reports real **eligibility**, not
successful input preflight or execution.

## Supported workbook subset

This is deliberately a conservative package profile, **not support for
arbitrary Excel-created `.xlsx` files**. The tested baseline is openpyxl
3.1.5-generated, default workbooks:

- One or more visible, plain worksheets with unique names.
- Unformatted cell values: ordinary text (including escaped XML characters),
  finite numbers, booleans, and blank cells. Duplicated values are allowed.
  Integer writes are limited to absolute values at most `10**308`; numeric
  storage is subject to normal Excel/openpyxl floating-point precision limits.
- Default openpyxl workbook, worksheet, stylesheet, theme, view, and page
  metadata. Metadata is not a fidelity guarantee; document timestamps may
  change on adapter save.
- Inline text strings only. Shared-string packages, including otherwise
  simple files from some Excel producers, are rejected in this phase.
- A maximum 100,000-cell bounding rectangle per worksheet (from A1), per
  requested read/write extent, and per checked range; standard Excel row/
  column boundaries and forward A1 ranges are enforced.
- At most 20 MB total uncompressed ZIP content. Cell text is limited to
  32,767 characters with no illegal XML control characters.

Empty worksheets and blank reads are supported. An empty string may round-trip
through openpyxl as a blank (`None`); exact empty-string-versus-blank fidelity
is not promised. Numeric storage does not preserve the Python distinction
between an integral float and an integer. No formula calculation is provided.

Before loading, the adapter checks allowed package parts and XML structures,
compares non-data settings against the default profile, and rejects unknown
or non-default features. It does not try to strip features to make a file fit.

Known rejected cases include `.xlsm`/VBA, formulas, date/time/error cells,
custom number formats/fonts/styles, merged cells, hidden sheets, protection,
custom layouts/views, hyperlinks, tables, charts, drawings/images, comments,
conditional formatting, data validation, named ranges, external links,
pivots, shared strings, custom XML, and unknown package parts/extensions.
This may reject a workbook that openpyxl itself can open. That is intentional:
opening and saving through openpyxl is not proof of feature preservation.

File-format preflight is read-only. Adapter callers must make a staged copy
before calling `write_values`/`save`; the runner never writes an input path.
Direct manipulation of the underlying openpyxl object is not a supported
adapter interface. Phase 2 real jobs contain validation only, so their final
workbook bytes are identical to the staged input copy.

## Execution, checks, and failure behavior

1. Schema validation resolves paths and validates references without checking
   input existence. Relative workbook paths are relative to the YAML file.
2. Real runtime capability preflight checks **every step**, then inspects all
   configured input workbooks. No unsupported operation is skipped.
3. Collision policy rejects any existing final output, including dry runs.
4. Normal runs create a temporary staging directory inside the output parent
   (creating the parent if needed). Dry runs use a system temporary directory
   and do not create the final output directory.
5. The base input is copied byte-for-byte to `staged-output.xlsx`; the staged
   package is checked again. Inputs/templates remain read-only.
6. All checks use the staged path, never a not-yet-published final path:
   - `file_exists`: staged file exists and is non-empty.
   - `sheet_exists`: named worksheet exists.
   - `range_nonempty`: at least one value is not blank/empty text; `0` and
     `false` count as present.
   - `range_equals`: scalar expected value for one cell, or a rectangular
     two-dimensional matrix exactly matching the range. Flat lists and
     mismatched shapes fail capability preflight. Values are never logged.
7. Only after all steps pass, a normal run atomically creates the final name
   using a same-filesystem hard link. An output created by another run in the
   meantime causes publication failure without overwriting it. No partial
   final copy is exposed. The staged name is subsequently removed.
8. Dry runs and fake runs never publish. Temporary workspaces are cleaned on
   success or processing failure. A failure before publication has no completed
   final output; an empty newly created output parent may remain. A cleanup
   failure after successful publication is reported as a filesystem error;
   the already-validated complete output is not deleted.

Publication requires filesystem hard-link support (for example NTFS or a
normal local POSIX filesystem). Unsupported/network filesystems may fail;
there is no unsafe copy-to-final fallback. This is no-overwrite publication,
not a distributed transaction or a power-loss durability guarantee. No email
service is present, including in real dry runs.

## Reproducible real CLI example

From the repository root in PowerShell:

```powershell
python -m pip install -e ".[dev]"
python -m excel_automation.cli generate-samples
python -m excel_automation.cli --config config\jobs.phase2.yaml validate
python -m excel_automation.cli --config config\jobs.phase2.yaml list-jobs
python -m excel_automation.cli --config config\jobs.phase2.yaml run plain-xlsx-check --dry-run
python -m excel_automation.cli --config config\jobs.phase2.yaml run plain-xlsx-check
python -m pytest -q
```

The input is generated at `examples/workbooks/phase2_sample.xlsx`; normal
success publishes `examples/output/phase2_result.xlsx`. The real run summary
starts with `REAL DIRECT_XLSX RUN SUCCEEDED`. Run dry-run before normal
publication: a later dry-run correctly fails the same collision check if the
final output already exists. Repeating sample generation also refuses to
overwrite its input fixture. Use a fresh location or explicitly move an old
artifact before rerunning.

## Portable test coverage

Tests generate non-sensitive workbooks in pytest temporary directories rather
than requiring production inputs or committed binary fixtures. They cover
staging byte equality, source SHA-256 preservation, value read/write round-trip,
empty/multiple worksheets, duplicate values, range bounds, malformed/missing
files, file/sheet/range checks, normal publication, real dry-run, all-step
capability rejection, known unsupported features/package parts, collision
races, publication errors, failure cleanup, safe error logs, and real CLI
execution. Phase 1 fake-mode tests remain in the suite. No VBA, desktop Excel,
rendering, or real email is exercised or claimed.
