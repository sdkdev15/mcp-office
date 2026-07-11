"""Auto-size engine — fits text into containers by binary-searching font size.

Uses Pillow to measure text dimensions and iteratively finds the optimal
font size that fits within a given bounding box.
"""

from __future__ import annotations

from typing import Optional

from PIL import Image, ImageDraw, ImageFont


class AutoSizeEngine:
    """Auto-fit text to container dimensions."""

    # Font cache
    _font_cache: dict[str, ImageFont.FreeTypeFont] = {}

    @classmethod
    def _get_font(cls, font_name: str, font_size: int, bold: bool = False) -> ImageFont.FreeTypeFont:
        """Load a font with caching."""
        cache_key = f"{font_name}:{font_size}:{bold}"
        if cache_key in cls._font_cache:
            return cls._font_cache[cache_key]

        # Try system fonts
        system_fonts = {
            "Calibri": "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
            "Arial": "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
            "Georgia": "/usr/share/fonts/truetype/dejavu/DejaVuSerif.ttf",
            "Times New Roman": "/usr/share/fonts/truetype/dejavu/DejaVuSerif.ttf",
            "Segoe UI": "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
            "Consolas": "/usr/share/fonts/truetype/dejavu/DejaVuSansMono.ttf",
            "monospace": "/usr/share/fonts/truetype/dejavu/DejaVuSansMono.ttf",
        }

        bold_paths = {
            "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf": "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
            "/usr/share/fonts/truetype/dejavu/DejaVuSerif.ttf": "/usr/share/fonts/truetype/dejavu/DejaVuSerif-Bold.ttf",
            "/usr/share/fonts/truetype/dejavu/DejaVuSansMono.ttf": "/usr/share/fonts/truetype/dejavu/DejaVuSansMono-Bold.ttf",
        }

        font_path = system_fonts.get(font_name, "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf")
        if bold and font_path in bold_paths:
            font_path = bold_paths[font_path]

        try:
            font = ImageFont.truetype(font_path, font_size)
        except (OSError, IOError):
            font = ImageFont.load_default()

        cls._font_cache[cache_key] = font
        return font

    @staticmethod
    def _measure_text(text: str, font: ImageFont.FreeTypeFont, max_width_px: int) -> tuple[int, int]:
        """Measure text dimensions, handling multi-line."""
        img = Image.new("RGBA", (1, 1), (0, 0, 0, 0))
        draw = ImageDraw.Draw(img)

        lines = text.split("\n")
        total_h = 0
        max_w = 0

        for line in lines:
            bbox = draw.textbbox((0, 0), line, font=font)
            line_w = bbox[2] - bbox[0]
            line_h = bbox[3] - bbox[1]
            max_w = max(max_w, line_w)
            total_h += line_h

        return max_w, total_h

    @classmethod
    def fit_text(
        cls,
        text: str,
        max_width_inches: float,
        max_height_inches: float,
        font_name: str = "Calibri",
        min_size: int = 6,
        max_size: int = 96,
        bold: bool = False,
        dpi: int = 96,
    ) -> int:
        """Binary-search for the largest font size that fits text in container.

        Args:
            text: Text to measure.
            max_width_inches: Container width in inches.
            max_height_inches: Container height in inches.
            font_name: Font name to use.
            min_size: Minimum font size to try.
            max_size: Maximum font size to try.
            bold: Whether to use bold variant.
            dpi: DPI for conversion.

        Returns:
            Optimal font size in points.
        """
        max_w_px = int(max_width_inches * dpi)
        max_h_px = int(max_height_inches * dpi)

        if not text:
            return max_size

        lo, hi = min_size, max_size
        best = min_size

        while lo <= hi:
            mid = (lo + hi) // 2
            font = cls._get_font(font_name, mid, bold=bold)
            w, h = cls._measure_text(text, font, max_w_px)

            if w <= max_w_px and h <= max_h_px:
                best = mid
                lo = mid + 1  # try larger
            else:
                hi = mid - 1  # too large, try smaller

        return best

    @classmethod
    def calculate_lines(
        cls,
        text: str,
        font_size: int,
        max_width_inches: float,
        font_name: str = "Calibri",
        dpi: int = 96,
    ) -> list[str]:
        """Split text into lines that fit within max_width."""
        max_w_px = int(max_width_inches * dpi)
        font = cls._get_font(font_name, font_size)
        img = Image.new("RGBA", (1, 1), (0, 0, 0, 0))
        draw = ImageDraw.Draw(img)

        words = text.split()
        if not words:
            return []

        lines = []
        current_line = words[0]

        for word in words[1:]:
            test_line = current_line + " " + word
            bbox = draw.textbbox((0, 0), test_line, font=font)
            w = bbox[2] - bbox[0]
            if w <= max_w_px:
                current_line = test_line
            else:
                lines.append(current_line)
                current_line = word

        lines.append(current_line)
        return lines


# ── Convenience Function ──

def auto_fit_text_size(
    text: str,
    max_width: float,
    max_height: float,
    font_name: str = "Calibri",
    min_size: int = 8,
    max_size: int = 72,
) -> int:
    """Quick auto-fit font size calculation."""
    return AutoSizeEngine.fit_text(
        text=text,
        max_width_inches=max_width,
        max_height_inches=max_height,
        font_name=font_name,
        min_size=min_size,
        max_size=max_size,
    )


def estimate_content_lines(text: str, font_size: int, max_width: float, font_name: str = "Calibri") -> list[str]:
    """Quick line splitting for content."""
    return AutoSizeEngine.calculate_lines(text, font_size, max_width, font_name)
