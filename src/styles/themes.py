"""Pre-built theme definitions for consistent document styling."""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Optional


@dataclass
class ThemeColors:
    """Color palette for a theme."""
    primary: str = "#2563EB"
    secondary: str = "#64748B"
    accent: str = "#F59E0B"
    background: str = "#FFFFFF"
    text: str = "#1E293B"
    text_light: str = "#64748B"
    success: str = "#10B981"
    warning: str = "#F59E0B"
    error: str = "#EF4444"
    header_bg: str = "#1E40AF"
    header_text: str = "#FFFFFF"
    table_alt_row: str = "#F1F5F9"
    border: str = "#E2E8F0"


@dataclass
class ThemeFonts:
    """Font configuration for a theme."""
    body: str = "Calibri"
    heading: str = "Calibri"
    monospace: str = "Consolas"
    body_size: int = 11
    heading1_size: int = 26
    heading2_size: int = 20
    heading3_size: int = 16
    heading4_size: int = 14


@dataclass
class Theme:
    """Complete theme definition."""
    name: str
    colors: ThemeColors = field(default_factory=ThemeColors)
    fonts: ThemeFonts = field(default_factory=ThemeFonts)
    description: str = ""


# ── Pre-built Themes ──

THEMES: dict[str, Theme] = {
    "corporate": Theme(
        name="corporate",
        colors=ThemeColors(
            primary="#1E40AF",
            secondary="#475569",
            accent="#F59E0B",
            background="#FFFFFF",
            text="#1E293B",
            text_light="#64748B",
            header_bg="#1E3A8A",
            header_text="#FFFFFF",
            table_alt_row="#F8FAFC",
            border="#CBD5E1",
        ),
        fonts=ThemeFonts(
            body="Calibri",
            heading="Calibri Light",
        ),
        description="Professional corporate style with blue tones, clean and formal",
    ),
    "minimal": Theme(
        name="minimal",
        colors=ThemeColors(
            primary="#000000",
            secondary="#666666",
            accent="#333333",
            background="#FFFFFF",
            text="#111111",
            text_light="#777777",
            header_bg="#000000",
            header_text="#FFFFFF",
            table_alt_row="#FAFAFA",
            border="#E5E5E5",
        ),
        fonts=ThemeFonts(
            body="Arial",
            heading="Arial",
        ),
        description="Clean minimal style with black and white, modern and simple",
    ),
    "creative": Theme(
        name="creative",
        colors=ThemeColors(
            primary="#7C3AED",
            secondary="#EC4899",
            accent="#FBBF24",
            background="#FAF5FF",
            text="#1F2937",
            text_light="#6B7280",
            header_bg="#7C3AED",
            header_text="#FFFFFF",
            table_alt_row="#F3E8FF",
            border="#E9D5FF",
        ),
        fonts=ThemeFonts(
            body="Segoe UI",
            heading="Segoe UI Light",
        ),
        description="Vibrant creative style with purple and pink gradients",
    ),
    "academic": Theme(
        name="academic",
        colors=ThemeColors(
            primary="#1E3A5F",
            secondary="#4A5568",
            accent="#2B6CB0",
            background="#FFFFFF",
            text="#2D3748",
            text_light="#718096",
            header_bg="#1A365D",
            header_text="#FFFFFF",
            table_alt_row="#EDF2F7",
            border="#CBD5E0",
        ),
        fonts=ThemeFonts(
            body="Times New Roman",
            heading="Times New Roman",
            body_size=12,
            heading1_size=24,
            heading2_size=18,
            heading3_size=16,
            heading4_size=14,
        ),
        description="Academic style with serif fonts, formal and traditional",
    ),
    "dark": Theme(
        name="dark",
        colors=ThemeColors(
            primary="#60A5FA",
            secondary="#94A3B8",
            accent="#FBBF24",
            background="#1E293B",
            text="#F1F5F9",
            text_light="#94A3B8",
            header_bg="#0F172A",
            header_text="#F1F5F9",
            table_alt_row="#334155",
            border="#475569",
        ),
        fonts=ThemeFonts(
            body="Segoe UI",
            heading="Segoe UI",
        ),
        description="Dark mode style with blue accents, easy on the eyes",
    ),
}


def get_theme(name: str) -> Theme:
    """Get a theme by name.

    Args:
        name: Theme name.

    Returns:
        Theme instance. Defaults to 'corporate' if not found.
    """
    return THEMES.get(name.lower(), THEMES["corporate"])


def list_themes() -> list[dict]:
    """List all available themes.

    Returns:
        List of theme information dictionaries.
    """
    return [
        {
            "name": theme.name,
            "description": theme.description,
            "primary_color": theme.colors.primary,
            "body_font": theme.fonts.body,
            "heading_font": theme.fonts.heading,
        }
        for theme in THEMES.values()
    ]


def get_available_theme_names() -> list[str]:
    """Get list of available theme names.

    Returns:
        List of theme name strings.
    """
    return list(THEMES.keys())