"""Icon renderer -- renders icon shapes to PNG images using Pillow.

Uses Pillow's ImageDraw for basic shapes and Image for alpha compositing.
Icons are drawn as simplified vector shapes (lines, circles, rects) rather
than full SVG paths, keeping dependencies minimal (Pillow only).

Cached by hash(icon_name + size + color + bg) to avoid redundant rendering.
"""

from __future__ import annotations

import hashlib
import os
from pathlib import Path
from typing import Optional

from PIL import Image, ImageDraw, ImageFont

from src.utils.logger import get_logger

log = get_logger("icon_renderer")

# Cache directory
_CACHE_DIR = os.path.join(os.path.dirname(__file__), "..", "..", ".icon_cache")


def _ensure_cache_dir() -> Path:
    """Create the icon cache directory if it doesn't exist."""
    cache_path = Path(_CACHE_DIR)
    cache_path.mkdir(parents=True, exist_ok=True)
    return cache_path


def _cache_key(icon_name: str, size: int, color: str, bg_color: Optional[str], bg_shape: str) -> str:
    """Generate a cache key for an icon render."""
    raw = f"{icon_name}:{size}:{color}:{bg_color}:{bg_shape}"
    return hashlib.md5(raw.encode()).hexdigest()


# ── Icon Shape Definitions ──
# Each icon is defined as a list of draw commands for Pillow ImageDraw.
# Commands: ("circle", cx, cy, r, fill), ("rect", x, y, w, h, fill),
#           ("line", x1, y1, x2, y2, fill, width), ("arc", ...)


def _get_icon_commands(icon_name: str) -> list[tuple]:
    """Get draw commands for an icon.

    Each command is a tuple of (shape_type, params...).
    All coordinates are in 0-24 range (matching SVG viewBox).
    """
    # Simplified icon definitions as Pillow-compatible shapes
    # Format: (type, x, y, w/h/r, fill) for shapes
    #          (type, x1, y1, x2, y2, fill, width) for lines

    icons = {
        # ── Security ──
        "lock": [
            ("rect", 7, 7, 10, 13, "fill"),      # body
            ("rect", 6, 14, 12, 7, "fill"),      # bottom
            ("circle", 12, 17, 2, "fill"),        # keyhole
            ("line", 12, 7, 12, 4, "fill", 2),    # shackle top
            ("arc_left", 9, 4, 6, 7, "fill", 2),  # shackle left
            ("arc_right", 15, 4, 18, 7, "fill", 2), # shackle right
        ],
        "unlock": [
            ("rect", 7, 11, 10, 10, "fill"),      # body
            ("rect", 6, 18, 12, 3, "fill"),       # bottom
            ("circle", 12, 16, 2, "fill"),         # keyhole
            ("line", 15, 11, 15, 4, "fill", 2),    # shackle (open)
            ("arc_right", 15, 4, 18, 7, "fill", 2),
        ],
        "shield": [
            ("polygon", 12, 2, 4, 5, 4, 12, 8, 18, 12, 20, 16, 18, 20, 12, 20, 5, "fill"),
        ],
        "key": [
            ("circle", 8, 8, 5, "fill"),           # head
            ("circle", 8, 8, 2, "bg"),              # hole (drawn in bg color)
            ("line", 11, 5, 20, 14, "fill", 2),     # shaft
            ("line", 18, 12, 18, 16, "fill", 2),    # tooth 1
            ("line", 20, 10, 20, 14, "fill", 2),    # tooth 2
        ],
        "fingerprint": [
            ("arc", 10, 6, 4, "fill", 1.5),
            ("arc", 10, 8, 5, "fill", 1.5),
            ("arc", 10, 10, 6, "fill", 1.5),
            ("arc", 10, 12, 7, "fill", 1.5),
        ],

        # ── Server / Infra ──
        "server": [
            ("rect", 3, 3, 18, 4, "fill"),
            ("rect", 3, 9, 18, 4, "fill"),
            ("rect", 3, 15, 18, 4, "fill"),
            ("circle", 6, 5, 1, "bg"),
            ("circle", 6, 11, 1, "bg"),
            ("circle", 6, 17, 1, "bg"),
        ],
        "database": [
            ("ellipse", 4, 2, 16, 6, "fill"),
            ("rect", 4, 5, 16, 4, "fill"),
            ("ellipse", 4, 12, 16, 6, "fill"),
        ],
        "cloud": [
            ("circle", 10, 12, 6, "fill"),
            ("circle", 16, 12, 4, "fill"),
            ("circle", 7, 14, 3, "fill"),
            ("rect", 6, 13, 14, 3, "fill"),
        ],
        "docker": [
            ("rect", 2, 8, 20, 10, "fill"),
            ("rect", 4, 6, 3, 3, "fill"),
            ("rect", 8, 6, 3, 3, "fill"),
            ("rect", 12, 6, 3, 3, "fill"),
        ],
        "network": [
            ("circle", 4, 6, 2, "fill"),
            ("circle", 20, 6, 2, "fill"),
            ("circle", 12, 12, 2, "fill"),
            ("circle", 4, 18, 2, "fill"),
            ("circle", 20, 18, 2, "fill"),
            ("line", 6, 6, 10, 12, "fill", 1.5),
            ("line", 14, 12, 18, 6, "fill", 1.5),
            ("line", 6, 12, 4, 16, "fill", 1.5),
            ("line", 14, 12, 18, 16, "fill", 1.5),
        ],
        "cpu": [
            ("rect", 6, 2, 12, 20, "fill"),
            ("rect", 8, 4, 8, 16, "bg"),
            ("line", 2, 6, 6, 6, "fill", 1.5),
            ("line", 2, 10, 6, 10, "fill", 1.5),
            ("line", 2, 14, 6, 14, "fill", 1.5),
            ("line", 18, 6, 22, 6, "fill", 1.5),
            ("line", 18, 10, 22, 10, "fill", 1.5),
            ("line", 18, 14, 22, 14, "fill", 1.5),
        ],
        "memory": [
            ("rect", 2, 4, 20, 16, "fill"),
            ("rect", 4, 2, 2, 4, "fill"),
            ("rect", 8, 2, 2, 4, "fill"),
            ("rect", 14, 2, 2, 4, "fill"),
            ("rect", 18, 2, 2, 4, "fill"),
            ("rect", 4, 18, 2, 4, "fill"),
            ("rect", 8, 18, 2, 4, "fill"),
            ("rect", 14, 18, 2, 4, "fill"),
            ("rect", 18, 18, 2, 4, "fill"),
        ],
        "disk": [
            ("circle", 12, 12, 10, "fill"),
            ("circle", 12, 12, 3, "bg"),
        ],

        # ── Business ──
        "dollar": [
            ("line", 12, 2, 12, 22, "fill", 2),
            ("arc", 14, 5, 3, "fill", 2),
            ("arc", 10, 9, 4, "fill", 2),
            ("arc", 14, 13, 3, "fill", 2),
            ("arc", 10, 17, 4, "fill", 2),
        ],
        "trend-up": [
            ("line", 2, 20, 22, 20, "fill", 2),
            ("line", 2, 2, 22, 20, "fill", 2.5),
            ("polygon", 14, 4, 22, 4, 22, 12, "fill"),
        ],
        "trend-down": [
            ("line", 2, 20, 22, 20, "fill", 2),
            ("line", 2, 20, 22, 2, "fill", 2.5),
            ("polygon", 14, 20, 22, 20, 22, 12, "fill"),
        ],
        "target": [
            ("circle", 12, 12, 10, "fill", True),
            ("circle", 12, 12, 6, "fill", True),
            ("circle", 12, 12, 2, "fill"),
        ],
        "award": [
            ("polygon", 12, 2, 15, 9, 22, 9, 17, 14, 19, 21, 12, 17, 5, 21, 7, 14, 2, 9, 9, 9, "fill"),
        ],
        "users": [
            ("circle", 7, 7, 4, "fill"),
            ("path_arc", 2, 17, 12, 14, "fill"),
            ("circle", 17, 8, 3, "fill"),
            ("path_arc", 13, 19, 22, 16, "fill"),
        ],
        "calendar": [
            ("rect", 3, 4, 18, 17, "fill"),
            ("rect", 5, 2, 2, 3, "fill"),
            ("rect", 17, 2, 2, 3, "fill"),
            ("line", 3, 9, 21, 9, "bg", 1.5),
        ],

        # ── Actions ──
        "alert": [
            ("polygon", 12, 3, 2, 21, 22, 21, "fill"),
            ("line", 12, 9, 12, 14, "bg", 2.5),
            ("circle", 12, 17.5, 1, "bg"),
        ],
        "warning": [
            ("polygon", 12, 2, 2, 22, 22, 22, "fill"),
            ("line", 12, 7, 12, 13, "bg", 2.5),
            ("circle", 12, 16, 1.2, "bg"),
        ],
        "check": [
            ("line", 4, 12, 9, 17, "fill", 3),
            ("line", 9, 17, 20, 6, "fill", 3),
        ],
        "xmark": [
            ("line", 6, 6, 18, 18, "fill", 3),
            ("line", 18, 6, 6, 18, "fill", 3),
        ],
        "info": [
            ("circle", 12, 12, 10, "fill"),
            ("line", 12, 10, 12, 17, "bg", 2.5),
            ("circle", 12, 6.5, 1.2, "bg"),
        ],
        "play": [
            ("polygon", 6, 4, 6, 20, 20, 12, "fill"),
        ],
        "pause": [
            ("rect", 5, 4, 4, 16, "fill"),
            ("rect", 15, 4, 4, 16, "fill"),
        ],
        "refresh": [
            ("arc", 16, 8, 5, "start", 0, "end", 270, "fill", 2),
            ("arrow", 18, 3, 14, 4, "fill"),
            ("arc", 8, 16, 5, "start", 90, "end", 360, "fill", 2),
            ("arrow", 6, 21, 10, 20, "fill"),
        ],
        "settings": [
            ("circle", 12, 12, 4, "fill"),
            ("circle", 12, 12, 7, "fill", True),
            ("line", 12, 3, 12, 5, "fill", 2),
            ("line", 12, 19, 12, 21, "fill", 2),
            ("line", 3, 12, 5, 12, "fill", 2),
            ("line", 19, 12, 21, 12, "fill", 2),
        ],
        "gear": [
            ("circle", 12, 12, 4, "fill"),
            ("circle", 12, 12, 3, "bg"),
            ("rect", 10.5, 3, 3, 3, "fill"),
            ("rect", 10.5, 18, 3, 3, "fill"),
            ("rect", 3, 10.5, 3, 3, "fill"),
            ("rect", 18, 10.5, 3, 3, "fill"),
        ],

        # ── Data ──
        "chart-bar": [
            ("line", 2, 22, 22, 22, "fill", 2),
            ("rect", 4, 12, 4, 10, "fill"),
            ("rect", 10, 8, 4, 14, "fill"),
            ("rect", 16, 4, 4, 18, "fill"),
        ],
        "chart-pie": [
            ("circle", 12, 12, 10, "fill"),
            ("wedge", 12, 12, 10, "start", 0, "end", 120, "bg"),
        ],
        "chart-line": [
            ("line", 2, 22, 22, 22, "fill", 2),
            ("line", 2, 2, 2, 22, "fill", 2),
            ("line", 4, 16, 8, 10, "fill", 2.5),
            ("line", 8, 10, 14, 14, "fill", 2.5),
            ("line", 14, 14, 20, 4, "fill", 2.5),
            ("circle", 4, 16, 2, "fill"),
            ("circle", 8, 10, 2, "fill"),
            ("circle", 14, 14, 2, "fill"),
            ("circle", 20, 4, 2, "fill"),
        ],
        "table": [
            ("rect", 2, 3, 20, 18, "fill", True),
            ("line", 2, 12, 22, 12, "fill", 1.5),
            ("line", 12, 3, 12, 21, "fill", 1.5),
        ],
        "file": [
            ("polygon", 4, 2, 14, 2, 20, 8, 20, 22, 4, 22, "fill"),
            ("polygon", 14, 2, 14, 8, 20, 8, "bg"),
        ],
        "folder": [
            ("polygon", 2, 6, 8, 6, 10, 8, 22, 8, 22, 20, 2, 20, "fill"),
        ],

        # ── Communication ──
        "mail": [
            ("rect", 2, 4, 20, 16, "fill"),
            ("line", 2, 4, 12, 12, "bg", 2),
            ("line", 22, 4, 12, 12, "bg", 2),
        ],
        "chat": [
            ("rect", 2, 4, 20, 12, "fill"),
            ("polygon", 6, 16, 2, 16, 4, 20, "fill"),
        ],
        "phone": [
            ("rect", 6, 2, 8, 20, "fill"),
            ("circle", 10, 5, 1, "bg"),
            ("rect", 8, 18, 4, 2, "bg"),
        ],
        "globe": [
            ("circle", 12, 12, 10, "fill"),
            ("ellipse", 8, 4, 8, 16, "bg", True),
            ("line", 4, 12, 20, 12, "bg", 1.5),
        ],

        # ── DevOps ──
        "git-branch": [
            ("circle", 6, 6, 3, "fill"),
            ("circle", 6, 18, 3, "fill"),
            ("circle", 18, 18, 3, "fill"),
            ("line", 6, 9, 6, 15, "fill", 2),
            ("line", 6, 9, 12, 9, "fill", 2),
            ("line", 12, 9, 12, 15, "fill", 2),
            ("line", 12, 15, 18, 15, "fill", 2),
        ],
        "terminal": [
            ("rect", 2, 3, 20, 18, "fill"),
            ("line", 5, 9, 9, 13, "bg", 2),
            ("line", 9, 13, 5, 17, "bg", 2),
            ("line", 13, 17, 19, 17, "bg", 2),
        ],
        "code": [
            ("line", 6, 6, 2, 12, "fill", 2.5),
            ("line", 2, 12, 6, 18, "fill", 2.5),
            ("line", 18, 6, 22, 12, "fill", 2.5),
            ("line", 22, 12, 18, 18, "fill", 2.5),
            ("line", 14, 4, 10, 20, "fill", 2),
        ],
        "pipeline": [
            ("rect", 2, 4, 5, 5, "fill"),
            ("rect", 9, 4, 5, 5, "fill"),
            ("rect", 16, 4, 5, 5, "fill"),
            ("rect", 2, 14, 5, 5, "fill"),
            ("rect", 9, 14, 5, 5, "fill"),
            ("rect", 16, 14, 5, 5, "fill"),
            ("line", 7, 6.5, 9, 6.5, "bg", 1.5),
            ("line", 14, 6.5, 16, 6.5, "bg", 1.5),
            ("line", 4.5, 9, 4.5, 14, "bg", 1.5),
            ("line", 11.5, 9, 11.5, 14, "bg", 1.5),
            ("line", 18.5, 9, 18.5, 14, "bg", 1.5),
        ],
        "bug": [
            ("circle", 12, 10, 5, "fill"),
            ("rect", 10, 15, 4, 6, "fill"),
            ("line", 7, 10, 2, 6, "fill", 1.5),
            ("line", 17, 10, 22, 6, "fill", 1.5),
            ("line", 6, 14, 1, 14, "fill", 1.5),
            ("line", 18, 14, 23, 14, "fill", 1.5),
            ("line", 7, 18, 2, 22, "fill", 1.5),
            ("line", 17, 18, 22, 22, "fill", 1.5),
        ],

        # ── Extended Icons ──
        "arrow-right": [
            ("line", 4, 12, 18, 12, "fill", 2.5),
            ("polygon", 14, 6, 20, 12, 14, 18, "fill"),
        ],
        "arrow-left": [
            ("line", 6, 12, 20, 12, "fill", 2.5),
            ("polygon", 10, 6, 4, 12, 10, 18, "fill"),
        ],
        "arrow-up": [
            ("line", 12, 6, 12, 20, "fill", 2.5),
            ("polygon", 6, 10, 12, 4, 18, 10, "fill"),
        ],
        "arrow-down": [
            ("line", 12, 4, 12, 18, "fill", 2.5),
            ("polygon", 6, 14, 12, 20, 18, 14, "fill"),
        ],
        "flow": [
            ("rect", 2, 3, 6, 6, "fill"),
            ("rect", 16, 3, 6, 6, "fill"),
            ("rect", 16, 15, 6, 6, "fill"),
            ("line", 8, 6, 16, 6, "fill", 2),
            ("line", 19, 9, 19, 15, "fill", 2),
        ],
        "connection": [
            ("line", 2, 12, 8, 12, "fill", 2.5),
            ("line", 16, 12, 22, 12, "fill", 2.5),
            ("polygon", 8, 8, 16, 12, 8, 16, "fill"),
        ],
        "clock": [
            ("circle", 12, 12, 10, "fill"),
            ("line", 12, 7, 12, 12, "bg", 2),
            ("line", 12, 12, 16, 14, "bg", 2),
        ],
        "hourglass": [
            ("polygon", 4, 2, 20, 2, 20, 4, 4, 4, "fill"),
            ("polygon", 4, 20, 20, 20, 20, 18, 4, 18, "fill"),
            ("polygon", 5, 5, 19, 5, 12, 12, "fill"),
            ("polygon", 12, 12, 19, 19, 5, 19, "fill"),
        ],
        "process": [
            ("rect", 3, 3, 7, 7, "fill"),
            ("rect", 14, 3, 7, 7, "fill"),
            ("rect", 3, 14, 7, 7, "fill"),
            ("rect", 14, 14, 7, 7, "fill"),
            ("line", 10, 6.5, 14, 6.5, "bg", 1.5),
            ("line", 10, 17.5, 14, 17.5, "bg", 1.5),
            ("line", 6.5, 10, 6.5, 14, "bg", 1.5),
            ("line", 17.5, 10, 17.5, 14, "bg", 1.5),
        ],
        "sync": [
            ("arc", 16, 9, 5, "start", 0, "end", 240, "fill", 2.5),
            ("polygon", 18, 3, 16, 6, 21, 6, "fill"),
            ("arc", 8, 15, 5, "start", 120, "end", 360, "fill", 2.5),
            ("polygon", 6, 21, 8, 18, 3, 18, "fill"),
        ],
        "repeat": [
            ("line", 3, 8, 11, 8, "fill", 2),
            ("polygon", 10, 5, 14, 8, 10, 11, "fill"),
            ("line", 21, 16, 13, 16, "fill", 2),
            ("polygon", 14, 13, 10, 16, 14, 19, "fill"),
        ],
        "quality": [
            ("polygon", 12, 1, 14.5, 8, 22, 8, 16, 13, 18, 21, 12, 16.5, 6, 21, 8, 13, 2, 8, 9.5, 8, "fill"),
        ],
        "speed": [
            ("circle", 12, 12, 10, "fill"),
            ("line", 12, 7, 15, 11, "bg", 2),
            ("circle", 12, 12, 2, "bg"),
        ],
        "bell": [
            ("polygon", 12, 2, 6, 6, 5, 13, 5, 15, 19, 15, 19, 13, 18, 6, "fill"),
            ("circle", 12, 18, 2, "fill"),
        ],
        "megaphone": [
            ("polygon", 2, 5, 14, 5, 18, 9, 14, 13, 2, 13, "fill"),
            ("rect", 1, 7, 2, 4, "fill"),
            ("rect", 15, 7, 5, 4, "fill"),
        ],
        "load-balancer": [
            ("circle", 12, 4, 2, "fill"),
            ("rect", 3, 8, 18, 3, "fill"),
            ("line", 6, 11, 6, 16, "fill", 2),
            ("line", 18, 11, 18, 16, "fill", 2),
            ("rect", 2, 16, 6, 4, "fill"),
            ("rect", 16, 16, 6, 4, "fill"),
        ],
        "firewall": [
            ("rect", 3, 3, 18, 18, "fill"),
            ("circle", 12, 12, 4, "bg"),
            ("line", 12, 8, 12, 4, "bg", 2),
        ],
        "backup": [
            ("circle", 12, 12, 10, "fill"),
            ("circle", 12, 12, 3, "bg"),
            ("polygon", 12, 15, 15, 19, 9, 19, "fill"),
        ],
        "monitor": [
            ("rect", 2, 3, 20, 12, "fill"),
            ("rect", 9, 15, 6, 3, "fill"),
            ("rect", 7, 18, 10, 2, "fill"),
        ],
        "web": [
            ("circle", 12, 12, 10, "fill"),
            ("ellipse", 7, 5, 10, 14, "bg", True),
            ("line", 3, 12, 21, 12, "bg", 1.5),
        ],
        "error": [
            ("circle", 12, 12, 10, "fill"),
            ("line", 7, 7, 17, 17, "bg", 2.5),
            ("line", 17, 7, 7, 17, "bg", 2.5),
        ],
        "success": [
            ("circle", 12, 12, 10, "fill"),
            ("line", 7, 12, 10, 15, "bg", 2.5),
            ("line", 10, 15, 17, 8, "bg", 2.5),
        ],
        "question": [
            ("circle", 12, 12, 10, "fill"),
            ("arc", 9, 7, 5, "start", 0, "end", 180, "bg", 2),
            ("circle", 12, 17, 1, "bg"),
        ],
        "presentation": [
            ("rect", 2, 2, 20, 14, "fill"),
            ("rect", 8, 16, 8, 2, "fill"),
            ("rect", 9, 18, 6, 3, "fill"),
        ],
        "image": [
            ("rect", 2, 3, 20, 18, "fill"),
            ("circle", 7, 8, 2, "bg"),
            ("polygon", 4, 12, 8, 8, 12, 14, 16, 10, 20, 13, 20, 20, 4, 20, "bg"),
        ],
        "api": [
            ("rect", 3, 5, 5, 14, "fill"),
            ("rect", 16, 5, 5, 14, "fill"),
            ("line", 8, 8, 16, 8, "fill", 2),
            ("polygon", 14, 5, 18, 8, 14, 11, "fill"),
            ("line", 16, 16, 8, 16, "fill", 2),
            ("polygon", 10, 13, 6, 16, 10, 19, "fill"),
        ],
        "upload": [
            ("rect", 3, 16, 18, 5, "fill"),
            ("line", 12, 4, 12, 14, "fill", 2.5),
            ("polygon", 7, 9, 12, 4, 17, 9, "fill"),
        ],
        "download": [
            ("rect", 3, 16, 18, 5, "fill"),
            ("line", 12, 4, 12, 14, "fill", 2.5),
            ("polygon", 7, 9, 12, 14, 17, 9, "fill"),
        ],
        "link": [
            ("arc", 7, 9, 5, "start", 45, "end", 225, "fill", 2.5),
            ("arc", 17, 9, 5, "start", -45, "end", 135, "fill", 2.5),
        ],
        "search": [
            ("circle", 11, 11, 8, "fill"),
            ("line", 17, 17, 21, 21, "fill", 2.5),
        ],
        "incident": [
            ("polygon", 12, 2, 2, 22, 22, 22, "fill"),
            ("line", 12, 7, 12, 13, "bg", 2.5),
            ("circle", 12, 16, 1.2, "bg"),
        ],
        "analytics": [
            ("line", 2, 20, 22, 20, "fill", 2),
            ("line", 2, 20, 2, 2, "fill", 2),
            ("rect", 4, 14, 4, 6, "fill"),
            ("rect", 10, 8, 4, 12, "fill"),
            ("rect", 16, 5, 4, 15, "fill"),
        ],
        "audit": [
            ("rect", 3, 3, 18, 18, "fill"),
            ("line", 7, 8, 17, 8, "bg", 1.5),
            ("line", 7, 12, 17, 12, "bg", 1.5),
            ("line", 7, 16, 13, 16, "bg", 1.5),
        ],
    }

    return icons.get(icon_name, [])


class IconRenderer:
    """Render icon shapes to PNG images using Pillow."""

    def __init__(self, cache_dir: Optional[str] = None):
        self.cache_dir = Path(cache_dir) if cache_dir else _ensure_cache_dir()

    def render(
        self,
        icon_name: str,
        size: int = 64,
        color: str = "#FFFFFF",
        bg_color: Optional[str] = None,
        bg_shape: str = "circle",
        padding: int = 8,
    ) -> str:
        """Render an icon to a PNG file.

        Args:
            icon_name: Name of the icon from icon_library.
            size: Output size in pixels (square).
            color: Icon fill color (hex string).
            bg_color: Background color (hex string or None for transparent).
            bg_shape: Background shape ("circle", "rounded_rect", or "none").
            padding: Padding in pixels inside the background shape.

        Returns:
            Path to the rendered PNG file.
        """
        key = _cache_key(icon_name, size, color, bg_color, bg_shape)
        cache_path = self.cache_dir / f"{key}.png"

        if cache_path.exists():
            return str(cache_path)

        # Create the image
        img_size = size
        img = Image.new("RGBA", (img_size, img_size), (0, 0, 0, 0))
        draw = ImageDraw.Draw(img)

        # Draw background shape
        if bg_color and bg_shape != "none":
            r, g, b = self._hex_to_rgb(bg_color)
            fill_color = (r, g, b, 255)

            if bg_shape == "circle":
                cx, cy = img_size // 2, img_size // 2
                radius = (img_size // 2) - 2
                draw.ellipse([cx - radius, cy - radius, cx + radius, cy + radius], fill=fill_color)
            elif bg_shape == "rounded_rect":
                margin = 4
                radius = 12
                draw.rounded_rect(
                    [margin, margin, img_size - margin, img_size - margin],
                    radius=radius,
                    fill=fill_color,
                )
            else:
                draw.rectangle([0, 0, img_size, img_size], fill=fill_color)

        # Get icon commands and scale to fit
        commands = _get_icon_commands(icon_name)
        if not commands:
            commands = [("circle", 12, 12, 5, "fill")]  # fallback

        # Convert hex color to RGB tuple
        r, g, b = self._hex_to_rgb(color)
        icon_color = (r, g, b, 255)
        bg_draw_color = (0, 0, 0, 0)  # transparent for "bg" type

        # Scale factor: 24 (icon coordinate space) -> img_size - 2*padding
        scale = (img_size - 2 * padding) / 24.0
        offset = padding

        for cmd in commands:
            cmd_type = cmd[0]
            fill_type = cmd[-1] if isinstance(cmd[-1], str) else None

            if fill_type == "bg":
                use_color = bg_draw_color
            else:
                use_color = icon_color

            if cmd_type == "circle":
                _, cx, cy, radius, *rest = cmd
                # Check for outline
                outline = False
                if isinstance(rest[0], bool):
                    outline = rest[0]
                scaled_cx = cx * scale + offset
                scaled_cy = cy * scale + offset
                scaled_r = radius * scale
                bbox = [scaled_cx - scaled_r, scaled_cy - scaled_r, scaled_cx + scaled_r, scaled_cy + scaled_r]
                if outline:
                    draw.ellipse(bbox, outline=use_color, width=max(1, int(2 * scale)))
                else:
                    draw.ellipse(bbox, fill=use_color)

            elif cmd_type == "rect":
                _, x, y, w, h, *rest = cmd
                outline = False
                if isinstance(rest[0], bool):
                    outline = rest[0]
                sx = x * scale + offset
                sy = y * scale + offset
                sw = w * scale
                sh = h * scale
                if outline:
                    draw.rectangle([sx, sy, sx + sw, sy + sh], outline=use_color, width=max(1, int(2 * scale)))
                else:
                    draw.rectangle([sx, sy, sx + sw, sy + sh], fill=use_color)

            elif cmd_type == "line":
                _, x1, y1, x2, y2, *rest = cmd
                width_idx = 5 if len(cmd) > 6 else 5
                line_width = 1.5
                if len(cmd) > 6:
                    line_width = cmd[6] if isinstance(cmd[6], (int, float)) else 1.5
                sx1 = x1 * scale + offset
                sy1 = y1 * scale + offset
                sx2 = x2 * scale + offset
                sy2 = y2 * scale + offset
                draw.line([sx1, sy1, sx2, sy2], fill=use_color, width=max(1, int(line_width * scale)))

            elif cmd_type == "polygon":
                _, *coords = cmd
                # Filter out non-numeric values (e.g., "fill" strings)
                numeric_coords = [c for c in coords if isinstance(c, (int, float))]
                points = [(c * scale + offset) for c in numeric_coords]
                draw.polygon(points, fill=use_color)

            elif cmd_type == "ellipse":
                _, x, y, w, h, *rest = cmd
                outline = False
                if rest and isinstance(rest[0], str) and rest[0] == "outline":
                    outline = True
                sx = x * scale + offset
                sy = y * scale + offset
                sw = w * scale
                sh = h * scale
                if outline:
                    draw.ellipse([sx, sy, sx + sw, sy + sh], outline=use_color, width=max(1, int(1.5 * scale)))
                else:
                    draw.ellipse([sx, sy, sx + sw, sy + sh], fill=use_color)

            elif cmd_type == "arc":
                # arc(cx, cy, radius, start=0, end=360, color, width)
                _, cx, cy, radius = cmd[1], cmd[2], cmd[3], cmd[4]
                start_angle = 0
                end_angle = 360
                i = 5
                while i < len(cmd) - 1:
                    if cmd[i] == "start":
                        start_angle = cmd[i + 1]
                    elif cmd[i] == "end":
                        end_angle = cmd[i + 1]
                    i += 2
                line_w = cmd[-1] if isinstance(cmd[-1], (int, float)) else 2
                scaled_cx = cx * scale + offset
                scaled_cy = cy * scale + offset
                scaled_r = radius * scale
                bbox = [scaled_cx - scaled_r, scaled_cy - scaled_r, scaled_cx + scaled_r, scaled_cy + scaled_r]
                draw.arc(bbox, start_angle, end_angle, fill=use_color, width=max(1, int(line_w * scale)))

            elif cmd_type in ("arc_left", "arc_right"):
                # Simplified arc drawing
                _, x, y, x2, y2, *rest = cmd
                line_w = cmd[5] if len(cmd) > 5 and isinstance(cmd[5], (int, float)) else 2
                sx = x * scale + offset
                sy = y * scale + offset
                sx2 = x2 * scale + offset
                sy2 = y2 * scale + offset
                draw.line([sx, sy, sx2, sy2], fill=use_color, width=max(1, int(line_w * scale)))

            elif cmd_type == "path_arc":
                _, x, y, w, h, *rest = cmd
                sx = x * scale + offset
                sy = y * scale + offset
                sw = w * scale
                sh = h * scale
                draw.ellipse([sx, sy - sh, sx + sw, sy], fill=use_color)

            elif cmd_type == "wedge":
                _, cx, cy, radius = cmd[1], cmd[2], cmd[3], cmd[4]
                start_angle = 0
                end_angle = 360
                i = 5
                while i < len(cmd) - 1:
                    if cmd[i] == "start":
                        start_angle = cmd[i + 1]
                    elif cmd[i] == "end":
                        end_angle = cmd[i + 1]
                    i += 2
                scaled_cx = cx * scale + offset
                scaled_cy = cy * scale + offset
                scaled_r = radius * scale
                bbox = [scaled_cx - scaled_r, scaled_cy - scaled_r, scaled_cx + scaled_r, scaled_cy + scaled_r]
                draw.pieslice(bbox, start_angle, end_angle, fill=use_color)

            elif cmd_type == "arrow":
                # arrow(x, y, x2, y2, color)
                _, ax, ay, ax2, ay2, *rest = cmd
                sax = ax * scale + offset
                say = ay * scale + offset
                sax2 = ax2 * scale + offset
                say2 = ay2 * scale + offset
                draw.line([sax, say, sax2, say2], fill=use_color, width=max(1, int(2 * scale)))
                # Simple arrowhead
                angle = ((say2 - say) / max(1, abs(sax2 - sax))) if sax2 != sax else 999
                head_size = 3 * scale
                if abs(sax2 - sax) > abs(say2 - say):
                    draw.polygon([
                        (sax2, say2),
                        (sax2 - head_size, say2 - head_size),
                        (sax2 - head_size, say2 + head_size),
                    ], fill=use_color)
                else:
                    draw.polygon([
                        (sax2, say2),
                        (sax2 - head_size, say2 - head_size),
                        (sax2 + head_size, say2 - head_size),
                    ], fill=use_color)

        # Save to cache
        cache_path.parent.mkdir(parents=True, exist_ok=True)
        img.save(str(cache_path), "PNG")
        return str(cache_path)

    def render_numbered_circle(
        self,
        number: int,
        size: int = 64,
        bg_color: str = "#1E40AF",
        text_color: str = "#FFFFFF",
    ) -> str:
        """Render a numbered circle (for agenda items).

        Args:
            number: Number to display.
            size: Output size in pixels.
            bg_color: Circle background color.
            text_color: Number text color.

        Returns:
            Path to the rendered PNG.
        """
        key = _cache_key(f"num_{number}", size, bg_color, text_color, "circle")
        cache_path = self.cache_dir / f"{key}.png"

        if cache_path.exists():
            return str(cache_path)

        img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
        draw = ImageDraw.Draw(img)

        # Draw circle
        r, g, b = self._hex_to_rgb(bg_color)
        cx, cy = size // 2, size // 2
        radius = (size // 2) - 2
        draw.ellipse([cx - radius, cy - radius, cx + radius, cy + radius], fill=(r, g, b, 255))

        # Draw number text
        text = str(number)
        try:
            font = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf", size - 20)
        except (OSError, IOError):
            try:
                font = ImageFont.truetype("/usr/share/fonts/TTF/DejaVuSans-Bold.ttf", size - 20)
            except (OSError, IOError):
                font = ImageFont.load_default()

        tr, tg, tb = self._hex_to_rgb(text_color)
        # Get text bounding box
        bbox = draw.textbbox((0, 0), text, font=font)
        tw = bbox[2] - bbox[0]
        th = bbox[3] - bbox[1]
        tx = cx - tw // 2 - bbox[0]
        ty = cy - th // 2 - bbox[1]
        draw.text((tx, ty), text, fill=(tr, tg, tb, 255), font=font)

        cache_path.parent.mkdir(parents=True, exist_ok=True)
        img.save(str(cache_path), "PNG")
        return str(cache_path)

    def render_batch(self, items: list[dict]) -> list[str]:
        """Render multiple icons at once.

        Each item is a dict with keys matching render() parameters.

        Returns:
            List of file paths.
        """
        paths = []
        for item in items:
            path = self.render(
                icon_name=item["icon_name"],
                size=item.get("size", 64),
                color=item.get("color", "#FFFFFF"),
                bg_color=item.get("bg_color"),
                bg_shape=item.get("bg_shape", "circle"),
                padding=item.get("padding", 8),
            )
            paths.append(path)
        return paths

    @staticmethod
    def _hex_to_rgb(hex_color: str) -> tuple[int, int, int]:
        """Convert hex color string to RGB tuple (supports #RGB and #RRGGBB)."""
        hex_color = hex_color.lstrip("#")
        if len(hex_color) == 3:
            hex_color = hex_color[0] * 2 + hex_color[1] * 2 + hex_color[2] * 2
        return (
            int(hex_color[0:2], 16),
            int(hex_color[2:4], 16),
            int(hex_color[4:6], 16),
        )