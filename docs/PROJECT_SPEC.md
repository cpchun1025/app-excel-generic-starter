# Excel Automation Application — Project Specification

## 1. Purpose

Build a maintainable, configuration-driven Python CLI application that
processes Excel workbooks, optionally runs VBA macros, produces output files,
and optionally delivers results by SMTP email.

Different jobs use different combinations of steps. Do not assume every job
needs Python data processing, a macro, or an email.

## 2. Initial scope

Build a minimal reusable framework first, demonstrated with several fake example jobs. Once it works and is tested,
generalize the parts that are genuinely shared by other jobs.

Until real job details are provided, use fake sample workbooks and clearly
marked example configuration. Do not guess production business rules.

## 3. Required workflow types

The application must eventually support:

- Python data preparation only.
- Excel VBA macro only.
- Python data preparation followed by an Excel VBA macro.
- Optional copying from source workbooks into destination sheets/ranges before
  macro execution.
- Optional output validation and SMTP email after processing.

The sequence of steps must be explicit and configurable for each job.

## 4. Job configuration

Support multiple named jobs. Each job may configure:

- Input workbook paths.
- Destination/template workbook path.
- Output directory and output filename.
- Source and destination sheets and ranges.
- Copy mode: values, values with formatting, or entire worksheet where feasible.
- Ordered processing steps.
- Python transformation/business-logic function, if applicable.
- Whether to run a macro; workbook, macro name, and arguments.
- Expected output files, sheets, and ranges to validate.
- Whether to send email and which email delivery formats to use.

Validate configuration before starting a job. Provide documented example
configuration with no real credentials or sensitive data.

## 5. Excel and macro handling

- State the operating-system and software prerequisites for each feature.
- For real VBA execution, assume Windows with Microsoft Excel installed.
- Do not use openpyxl as a VBA execution engine.
- Select direct workbook editing or desktop Excel automation according to
  workbook requirements. Avoid silently damaging VBA, charts, formatting, or
  other existing workbook features.
- Address the intended workbook and macro explicitly; do not depend on
  whichever Excel window or workbook happens to be active.
- Work on output copies by default. Protect source/template workbooks.
- Handle missing files, sheets, ranges, macros, locked files, and Excel errors.
- Close workbooks and release automation resources on both success and failure.
- Do not globally weaken Excel macro security or run untrusted macros.
- Document limitations and prerequisites of unattended Excel automation.

## 6. Email delivery

For each job, support applicable combinations of:

- Render a selected sheet/range as an HTML table in the email body.
- Render a selected sheet/range as an image and attach it.
- Export selected sheet/range content to an HTML file and attach it.
- Attach the finished Excel workbook.

Requirements:

- SMTP server, port, TLS mode, sender, recipients, subject, and body template
  must be configurable.
- Read SMTP credentials from environment variables or another approved secret
  source. Never commit or log them.
- Use correct MIME types and escape spreadsheet content in generated HTML.
- Clearly document how Excel-to-image rendering works and its limitations.
- Do not send email if a required processing or validation step fails.
- Provide --dry-run: process and validate, but never send email.
- Choose and document a method to reduce duplicate emails after retries.

## 7. Logging and error handling

- Write rotating log files to logs/ under the application/repository directory,
  regardless of the current working directory.
- Log timestamps, job name, run ID, step name, success/failure, and useful
  exception details.
- Show concise progress and errors in the console.
- Never log credentials or sensitive spreadsheet cell contents.
- Return meaningful CLI exit codes and a final run summary.
- Retry only operations known to be safe to repeat. Do not blindly rerun a
  macro or resend an email.

## 8. Testing

- Use pytest for Python transformations and business logic.
- Cover normal, empty, malformed, duplicate, and boundary-case input where
  relevant.
- Generate small, fake .xlsx workbooks for automated tests.
- Test configuration validation, copying, output validation, HTML generation,
  and email message construction.
- Mock Excel macro execution and SMTP delivery in unit tests.
- Provide a separate real-Excel integration test using a user-supplied .xlsm
  workbook. Do not claim a generated .xlsx contains a working VBA macro.

## 9. Expected project deliverables

- Python application source code with separated modules for configuration,
  orchestration, workbook operations, macro execution, rendering, and email.
- Example job configuration and fake sample workbooks.
- Automated tests and a separate integration-test procedure.
- Dependency declarations.
- README with installation, prerequisites, security notes, troubleshooting,
  limitations, and CLI examples.
- Commands to list jobs, run one job, run a dry run, generate sample files,
  and run tests.

## 10. Production job onboarding checklist

Fill these in before implementing a production workflow:

- Job name:
- Where it will run (Windows version / Excel available):
- Input workbook(s) and file types:
- Template/destination workbook and file type:
- Source sheets and ranges:
- Destination sheets and ranges:
- Copy mode:
- Python business rules, if any:
- Macro workbook and exact macro name, if any:
- Macro arguments, if any:
- Expected output file and validation checks:
- Email recipients and delivery format:
- SMTP provider and approved secret-storage method:
- Scheduling requirements:
- Sensitive-data restrictions:

Do not place real passwords or confidential workbook contents in this document.