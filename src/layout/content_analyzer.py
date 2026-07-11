"""Content analyzer — estimates content size, suggests layouts, splits content."""

from __future__ import annotations

from typing import Optional

from PIL import Image, ImageDraw, ImageFont


class ContentAnalyzer:
    """Analyze content and suggest optimal layouts."""

    @staticmethod
    def _get_system_font(font_name: str, size: int, bold: bool = False) -> ImageFont.FreeTypeFont:
        """Get a system font by name or fallback."""
        font_map = {
            "Calibri": "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
            "Calibri Light": "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
            "Arial": "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
            "Georgia": "/usr/share/fonts/truetype/dejavu/DejaVuSerif.ttf",
            "Times New Roman": "/usr/share/fonts/truetype/dejavu/DejaVuSerif.ttf",
            "Segoe UI": "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        }
        bold_map = {
            "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf": "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
            "/usr/share/fonts/truetype/dejavu/DejaVuSerif.ttf": "/usr/share/fonts/truetype/dejavu/DejaVuSerif-Bold.ttf",
        }
        path = font_map.get(font_name, "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf")
        if bold and path in bold_map:
            path = bold_map[path]
        try:
            return ImageFont.truetype(path, size)
        except (OSError, IOError):
            return ImageFont.load_default()

    @staticmethod
    def estimate_text_height(
        text: str,
        font_size: int,
        max_width_inches: float,
        font_name: str = "Calibri",
        line_spacing: float = 1.2,
        dpi: int = 96,
    ) -> float:
        """Estimate total text height in inches.

        Args:
            text: Text content.
            font_size: Font size in points.
            max_width_inches: Container width in inches.
            font_name: Font name.
            line_spacing: Multiplier for line height.
            dpi: Display DPI.

        Returns:
            Estimated height in inches.
        """
        max_w_px = int(max_width_inches * dpi)
        font = ContentAnalyzer._get_system_font(font_name, font_size)
        img = Image.new("RGBA", (1, 1), (0, 0, 0, 0))
        draw = ImageDraw.Draw(img)

        lines = text.split("\n")
        wrapped_lines = []

        for line in lines:
            words = line.split()
            if not words:
                wrapped_lines.append("")
                continue
            current = words[0]
            for word in words[1:]:
                test = current + " " + word
                bbox = draw.textbbox((0, 0), test, font=font)
                w = bbox[2] - bbox[0]
                if w <= max_w_px:
                    current = test
                else:
                    wrapped_lines.append(current)
                    current = word
            wrapped_lines.append(current)

        # Measure a single line height
        if wrapped_lines:
            sample_bbox = draw.textbbox((0, 0), "Ag", font=font)
            line_h_px = (sample_bbox[3] - sample_bbox[1]) * line_spacing
            total_h_px = line_h_px * len(wrapped_lines)
            return total_h_px / dpi
        return 0.0

    @staticmethod
    def estimate_content_height(
        title: str,
        bullets: Optional[list[str]] = None,
        content: str = "",
        num_stats: int = 0,
        has_table: bool = False,
        table_rows: int = 0,
        font_name: str = "Calibri",
        title_size: int = 28,
        body_size: int = 18,
        max_width: float = 11.0,
    ) -> float:
        """Estimate total slide content height.

        Returns estimated height in inches so the layout engine can
        decide if content fits or needs adjustment.
        """
        total = 0.0

        # Title
        if title:
            total += ContentAnalyzer.estimate_text_height(
                title, title_size, max_width, font_name
            ) + 0.3

        # Bullets
        if bullets:
            bullet_text = "\n".join(f"• {b}" for b in bullets)
            total += ContentAnalyzer.estimate_text_height(
                bullet_text, body_size, max_width * 0.9, font_name
            ) + 0.2

        # Content paragraph
        if content:
            total += ContentAnalyzer.estimate_text_height(
                content, body_size, max_width, font_name
            ) + 0.2

        # Stats (assume stat boxes ~1.2 inches each)
        if num_stats:
            total += 1.8  # stat box height

        # Table
        if has_table:
            total += 0.5 + table_rows * 0.4

        return total

    @staticmethod
    def suggest_layout_for_content(
        title: str,
        bullets: Optional[list[str]] = None,
        content: str = "",
        num_stats: int = 0,
        has_table: bool = False,
        has_chart: bool = False,
        num_images: int = 0,
    ) -> str:
        """Suggest the optimal slide layout based on content type.

        Returns:
            Layout name: 'title_only', 'title_and_content', 'two_content',
            'comparison', 'blank', 'section_header', or 'visual_type'.
        """
        if not title and not bullets and not content:
            return "blank"

        if title and not bullets and not content and not has_table and not has_chart:
            return "section_header"

        if has_chart or num_images > 0:
            return "title_and_content"

        if has_table and bullets:
            return "two_content"

        if num_stats > 0 and num_stats <= 4:
            return "exec_summary"

        if bullets and len(bullets) <= 4 and not content:
            return "title_and_content"

        if content and bullets:
            return "two_content"

        return "title_and_content"

    @staticmethod
    def split_content_into_slides(
        title: str,
        bullets: list[str],
        max_bullets_per_slide: int = 5,
    ) -> list[dict]:
        """Split large content into multiple slides.

        Returns:
            List of slide dicts with 'title' and 'bullets' keys.
        """
        if len(bullets) <= max_bullets_per_slide:
            return [{"title": title, "bullets": bullets}]

        slides = []
        base_title = title.rstrip(".!?")
        for i in range(0, len(bullets), max_bullets_per_slide):
            chunk = bullets[i:i + max_bullets_per_slide]
            slide_title = f"{base_title} ({i // max_bullets_per_slide + 1})" if i > 0 else base_title
            slides.append({"title": slide_title, "bullets": chunk})

        return slides


# ── Convenience Functions ──

def estimate_slide_fit(title: str, bullets: list[str], max_slide_height: float = 5.5) -> tuple[bool, float]:
    """Check if content fits on one slide.

    Returns:
        (fits, estimated_height)
    """
    h = ContentAnalyzer.estimate_content_height(title=title, bullets=bullets)
    return h <= max_slide_height, h

def suggest_slides_from_data(data: dict) -> list[dict]:
    """Suggest a full slide deck structure from a data dict."""
    slides = []

    if data.get("title"):
        slides.append({"visual_type": "cover", "title": data["title"], "subtitle": data.get("subtitle", "")})

    if data.get("stats"):
        slides.append({"visual_type": "exec_summary", "title": "Key Metrics", "stats": data["stats"]})

    if data.get("sections"):
        for section in data["sections"]:
            slides.append({"visual_type": "section_header", "title": section.get("title", "")})
            if section.get("bullets"):
                slides.append({"layout": "title_and_content", "title": section["title"], "bullets": section["bullets"]})
            if section.get("table"):
                slides.append({"layout": "title_and_content", "title": section["title"], "table": section["table"]})

    if data.get("timeline"):
        slides.append({"visual_type": "timeline", "title": data.get("timeline_title", "Timeline"), "events": data["timeline"]})

    if data.get("comparison"):
        slides.append({"visual_type": "comparison", "title": data.get("comparison_title", "Comparison"), **data["comparison"]})

    if data.get("action_items"):
        slides.append({"visual_type": "cta", "title": "Next Steps", "content": "\n".join(f"• {a}" for a in data["action_items"])})

    return slides
