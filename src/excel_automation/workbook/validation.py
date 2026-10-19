"""Real Phase 2 checks operate only on the staged workbook."""

from pathlib import Path
from typing import Any

from excel_automation.errors import ExecutionError
from excel_automation.workbook.direct_xlsx import DirectXlsx, bounds

CHECK_TYPES = frozenset({"file_exists", "sheet_exists", "range_nonempty", "range_equals"})


def check_definition(check: dict[str, Any]) -> None:
    if check["type"] not in CHECK_TYPES:
        raise ExecutionError("Unsupported direct_xlsx validation check")
    if "range" in check:
        left, top, right, bottom = bounds(check["range"])
        if check["type"] == "range_equals":
            expected = check["expected"]
            if not isinstance(expected, list):
                expected = [[expected]]
            if (
                not expected
                or any(not isinstance(row, list) for row in expected)
                or len(expected) != bottom - top + 1
                or any(len(row) != right - left + 1 for row in expected)
            ):
                raise ExecutionError(
                    "range_equals expects a scalar for one cell or a rectangular matrix matching the range"
                )


def validate_output(path: Path, checks: list[dict[str, Any]]) -> None:
    for check in checks:
        check_definition(check)
        if not path.is_file() or path.stat().st_size == 0:
            raise ExecutionError("Staged output file is missing or empty")
        kind = check["type"]
        if kind == "file_exists":
            continue
        with DirectXlsx(path) as adapter:
            if not adapter.has_sheet(check["sheet"]):
                raise ExecutionError("Output validation failed: required worksheet is missing")
            if kind == "sheet_exists":
                continue
            values = adapter.read_values(check["sheet"], check["range"])
            if kind == "range_nonempty":
                if not any(value is not None and value != "" for row in values for value in row):
                    raise ExecutionError("Output validation failed: required range is empty")
            elif kind == "range_equals":
                expected = check["expected"]
                if not isinstance(expected, list):
                    expected = [[expected]]
                if any(
                    actual != wanted
                    or (isinstance(actual, bool) != isinstance(wanted, bool))
                    for actual_row, expected_row in zip(values, expected)
                    for actual, wanted in zip(actual_row, expected_row)
                ):
                    raise ExecutionError("Output validation failed: range did not match; values omitted")
