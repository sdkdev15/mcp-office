"""Structured JSON logging for production monitoring."""

from __future__ import annotations

import json
import sys
from pathlib import Path
from typing import Any

from loguru import logger

# Configuration guard — ensure handlers are only added once
_configured = False


class JSONFormatter:
    """Custom formatter that outputs structured JSON logs."""

    def format(self, record: dict[str, Any]) -> str:
        log_entry = {
            "level": record["level"].name,
            "message": record["message"],
            "time": record["time"].isoformat(),
            "module": record["name"],
        }
        if "exception" in record and record["exception"]:
            log_entry["exception"] = str(record["exception"])
        return json.dumps(log_entry)


def _configure_once() -> None:
    """Configure loguru handlers once at first import."""
    global _configured
    if _configured:
        return
    _configured = True

    # Remove default handler
    logger.remove()

    # Add structured logging to stdout
    logger.add(
        sys.stdout,
        format="{time:YYYY-MM-DD HH:mm:ss} | {level:<8} | {name}:{function}:{line} | {message}",
        level="INFO",
    )

    # Add error logging to stderr
    logger.add(
        sys.stderr,
        format="{time:YYYY-MM-DD HH:mm:ss} | {level:<8} | {message}",
        level="ERROR",
    )


# Configure on module import
_configure_once()


def get_logger(name: str = "mcp-office") -> logger:
    """Get a configured loguru logger instance.

    Args:
        name: Logger name prefix.

    Returns:
        Configured logger instance.
    """
    return logger.bind(name=name)


# Module-level logger convenience
log = get_logger()