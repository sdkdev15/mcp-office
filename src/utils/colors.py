"""Shared color conversion utilities."""

from __future__ import annotations


def hex_to_rgbcolor_tuple(hex_color: str) -> tuple[int, int, int]:
    """Convert hex color string to (r, g, b) tuple.

    Args:
        hex_color: Hex color string (e.g., '1E40AF' or '#1E40AF').

    Returns:
        Tuple of (r, g, b) integers.
    """
    c = hex_color.lstrip("#")
    r = int(c[0:2], 16)
    g = int(c[2:4], 16)
    b = int(c[4:6], 16)
    return (r, g, b)


def ensure_argb_hex(color: str) -> str:
    """Ensure color is a valid ARGB hex string for openpyxl (8 chars with alpha).

    Args:
        color: Hex color string.

    Returns:
        8-character ARGB hex string.
    """
    if not color:
        return "FF000000"
    c = color.lstrip("#")
    if len(c) == 8:
        return c
    if len(c) == 6:
        return "FF" + c
    return "FF" + c.ljust(6, "0")
