"""Small immutable models used after strict configuration validation."""

from dataclasses import dataclass
from pathlib import Path
from typing import Any


@dataclass(frozen=True)
class Step:
    id: str
    type: str
    definition: dict[str, Any]


@dataclass(frozen=True)
class Job:
    id: str
    inputs: dict[str, Path]
    base: str
    output: Path
    adapter: str
    steps: tuple[Step, ...]
    email: dict[str, Any] | None


@dataclass(frozen=True)
class Configuration:
    source_path: Path
    jobs: dict[str, Job]
