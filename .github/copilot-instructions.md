# Copilot instructions

This repository contains a configuration-driven Python application for Excel
processing, optional VBA macro execution, and SMTP email delivery.

## Source of truth

- Read [the project specification](../docs/PROJECT_SPEC.md) before planning or
  implementing application features.
- Follow the specification when requirements are clear. If requirements are
  missing or conflicting, state the assumption or ask before implementing.
- Do not invent sheet names, ranges, macro names, SMTP settings, or business
  rules for production jobs.

## Implementation rules

- Build reusable components; keep job-specific settings in configuration.
- Support Python-only, macro-only, and Python-then-macro workflows.
- Treat VBA execution as a Windows + installed desktop Excel feature.
- Never claim that openpyxl can execute VBA macros.
- Work on copies of source workbooks unless a job explicitly requires otherwise.
- Validate outputs before preparing or sending email.
- Never send a real email in tests or dry-run mode.
- Keep credentials out of source code, committed configuration, and logs.
- Save application log files under the repository's logs/ directory.
- Add pytest tests for Python transformations, business logic, configuration,
  and email construction. Mock Excel and SMTP in unit tests.
- Use fake, non-sensitive workbooks for examples and automated tests.

## How to work

- This project is a reusable framework for many Excel automation jobs, not
  an implementation of one production job.
- Do not require real workbook details before building the framework.
- Use fake sample workbooks and example job configurations to demonstrate
  supported workflows.
- Keep workflow selection, step order, workbook mappings, macro names, output
  validation, and email options configurable.
- Put complex job-specific Python business logic in testable plugins referenced
  by configuration; do not execute arbitrary code from configuration.
- Do not invent or implement production business rules.
- Make small, reviewable changes rather than generating the entire application
  in one step.
- Explain assumptions, changed files, and how to run relevant tests.
- Run relevant tests when possible. Never say tests passed unless they ran.