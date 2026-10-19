"""Safe YAML loading with duplicate-key rejection."""

from pathlib import Path
from typing import Any

import yaml
from yaml.constructor import ConstructorError

from excel_automation.errors import ConfigurationError


class _UniqueKeyLoader(yaml.SafeLoader):
    """SafeLoader that refuses ambiguous mappings."""


def _construct_mapping(
    loader: _UniqueKeyLoader, node: yaml.MappingNode, deep: bool = False
) -> dict[Any, Any]:
    mapping: dict[Any, Any] = {}
    for key_node, value_node in node.value:
        key = loader.construct_object(key_node, deep=deep)
        try:
            duplicate = key in mapping
        except TypeError as exc:
            raise ConstructorError(
                "while constructing a mapping",
                node.start_mark,
                "found an unhashable mapping key",
                key_node.start_mark,
            ) from exc
        if duplicate:
            raise ConstructorError(
                "while constructing a mapping",
                node.start_mark,
                f"found duplicate key {key!r}",
                key_node.start_mark,
            )
        mapping[key] = loader.construct_object(value_node, deep=deep)
    return mapping


_UniqueKeyLoader.add_constructor(
    yaml.resolver.BaseResolver.DEFAULT_MAPPING_TAG, _construct_mapping
)


def load_config(path: Path) -> tuple[dict[str, Any], Path]:
    """Load a UTF-8 YAML mapping and return it with its resolved source path."""
    config_path = path.expanduser().resolve()
    try:
        with config_path.open("r", encoding="utf-8") as stream:
            document = yaml.load(stream, Loader=_UniqueKeyLoader)
    except OSError as exc:
        raise ConfigurationError(
            f"Cannot read configuration file {config_path}: {exc.strerror or type(exc).__name__}"
        ) from exc
    except yaml.YAMLError as exc:
        mark = getattr(exc, "problem_mark", None)
        location = f" at line {mark.line + 1}, column {mark.column + 1}" if mark else ""
        raise ConfigurationError(
            f"Invalid YAML{location} ({type(exc).__name__})"
        ) from exc

    if not isinstance(document, dict):
        raise ConfigurationError("Configuration root must be a mapping")
    return document, config_path
