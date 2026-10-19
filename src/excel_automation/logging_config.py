"""Safe console and repository-root rotating file logging."""

import logging
from logging.handlers import RotatingFileHandler
from pathlib import Path


def repository_root() -> Path:
    return Path(__file__).resolve().parents[2]


def configure_logging(*, console: bool = True) -> logging.Logger:
    """Configure application logs independent of the process working directory."""
    log_dir = repository_root() / "logs"
    log_dir.mkdir(parents=True, exist_ok=True)

    logger = logging.getLogger("excel_automation")
    logger.setLevel(logging.INFO)
    logger.propagate = False
    for handler in logger.handlers[:]:
        logger.removeHandler(handler)
        handler.close()

    formatter = logging.Formatter(
        "%(asctime)s %(levelname)s %(name)s %(message)s",
        datefmt="%Y-%m-%dT%H:%M:%S",
    )
    file_handler = RotatingFileHandler(
        log_dir / "excel_automation.log",
        maxBytes=2_000_000,
        backupCount=5,
        encoding="utf-8",
    )
    file_handler.setFormatter(formatter)
    logger.addHandler(file_handler)
    if console:
        console_handler = logging.StreamHandler()
        console_handler.setFormatter(formatter)
        logger.addHandler(console_handler)
    return logger
