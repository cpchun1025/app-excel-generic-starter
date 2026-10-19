"""Errors surfaced by the framework's configuration and execution layers."""


class AutomationError(Exception):
    """Base class for expected framework errors."""


class ConfigurationError(AutomationError):
    """Raised when configuration cannot be loaded or validated."""


class ExecutionError(AutomationError):
    """Raised when a job cannot be executed by the selected runtime."""


class OutputCollisionError(ExecutionError):
    """Raised when a configured output path already exists."""
