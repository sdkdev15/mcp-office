"""Shared formatting utilities."""

from __future__ import annotations


def human_readable_size(size_bytes: int) -> str:
    """Convert bytes to human-readable size string.

    Args:
        size_bytes: Size in bytes.

    Returns:
        Human-readable size string (e.g., '1.5 MB').
    """
    for unit in ["B", "KB", "MB", "GB"]:
        if size_bytes < 1024:
            return f"{size_bytes:.1f} {unit}"
        size_bytes /= 1024
    return f"{size_bytes:.1f} TB"
