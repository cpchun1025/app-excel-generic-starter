"""Stage byte-for-byte copies; publish complete files without overwriting."""

import os
import shutil
from pathlib import Path

from excel_automation.errors import ExecutionError, OutputCollisionError


def stage_copy(source: Path, destination: Path) -> None:
    if source.resolve() == destination.resolve():
        raise ExecutionError("Staging destination must differ from source")
    try:
        with source.open("rb") as reader, destination.open("xb") as writer:
            shutil.copyfileobj(reader, writer)
    except OSError as exc:
        raise ExecutionError(f"Cannot stage workbook ({type(exc).__name__})") from exc


def publish_output(staged: Path, output: Path) -> None:
    """Atomically create a final name for a complete staged file, never overwrite."""
    try:
        os.link(staged, output)
    except FileExistsError as exc:
        raise OutputCollisionError("Output path already exists; publication refused") from exc
    except OSError as exc:
        raise ExecutionError(
            f"Cannot publish workbook ({type(exc).__name__}); "
            "publication requires same-filesystem hard-link support"
        ) from exc
