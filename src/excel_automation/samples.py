"""Generate a small non-sensitive Phase 2 workbook, never a VBA fixture."""

from pathlib import Path
from tempfile import TemporaryDirectory

from openpyxl import Workbook

from excel_automation.errors import ExecutionError
from excel_automation.workbook.copies import publish_output


def generate_sample(output: Path) -> None:
    if output.suffix.lower() != ".xlsx":
        raise ExecutionError("Sample generator creates .xlsx files only")
    output.parent.mkdir(parents=True, exist_ok=True)
    with TemporaryDirectory(prefix="excel-sample-", dir=output.parent) as directory:
        staged = Path(directory) / "sample.xlsx"
        workbook = Workbook()
        try:
            sheet = workbook.active
            sheet.title = "Sample"
            for row in [
                ["Item", "Quantity", "Active"],
                ["Fake pencils", 12, True],
                ["Fake notebooks", 5, False],
            ]:
                sheet.append(row)
            workbook.save(staged)
        finally:
            workbook.close()
        publish_output(staged, output)
