# Excel Automation Framework

Configuration-driven Python framework under incremental development.

## Current capabilities

- Strict version 1 YAML validation, job listing, rotating repository-root logs,
  and explicitly labeled fake-service simulation.
- Phase 2 real staging and validation for a **conservative plain `.xlsx`
  profile**, with no-overwrite publication and temporary-only real dry-run.
- Adapter-level basic value reads/writes, tested on generated non-sensitive
  workbook fixtures.

No configured copy mappings, callable plugins, VBA execution, rendering, or
SMTP exist yet. Arbitrary Excel workbook fidelity is not claimed. See
[DIRECT_XLSX.md](docs/DIRECT_XLSX.md) for supported files, rejected features,
filesystem requirements, and the real CLI example.

## Install and test

Python 3.11 or later, from the repository root:

```powershell
python -m pip install -e ".[dev]"
python -m pytest -q
python -m excel_automation.cli generate-samples
python -m excel_automation.cli --config config\jobs.phase2.yaml run plain-xlsx-check --dry-run
python -m excel_automation.cli --config config\jobs.phase2.yaml run plain-xlsx-check
```

The normal run publishes the validated generated input copy at
`examples/output/phase2_result.xlsx`. Existing outputs cause failure in both
normal and dry-run mode. Sample generation also refuses to overwrite files.
No desktop Excel, VBA fixture, network, or email credentials are needed.

## Design and limitations

- [Architecture](docs/ARCHITECTURE.md)
- [Configuration contract and status](docs/CONFIGURATION.md)
- [Phased implementation plan](docs/IMPLEMENTATION_PLAN.md)
- [Original fake-only example jobs](config/jobs.example.yaml)
- [Real Phase 2 example job](config/jobs.phase2.yaml)

Relative workbook paths resolve against the YAML directory. Logs go under
repository `logs/`, independent of the working directory. Spreadsheet cell
contents and credentials must not be logged. Never use untrusted macros;
future VBA automation will require Windows and installed desktop Excel.
