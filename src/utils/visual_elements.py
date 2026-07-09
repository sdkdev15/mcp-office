"""Visual elements generator -- creates composite visual elements as PNG images.

Generates: stat boxes, timeline items, flow diagrams, comparison cards,
and other premium visual elements suitable for PowerPoint slides and Word docs.
"""

from __future__ import annotations

import os
from pathlib import Path
from typing import Optional

from PIL import Image, ImageDraw, ImageFont

from src.styles.themes import get_theme, Theme
from src.utils.logger import get_logger
from src.utils.icon_renderer import IconRenderer, _cache_key, _ensure_cache_dir

log = get_logger("visual_elements")


class VisualElements:
    """Generate visual elements as PNG images for slides/documents."""

    def __init__(self, cache_dir: Optional[str] = None):
        self.cache_dir = Path(cache_dir) if cache_dir else _ensure_cache_dir()
        self.icon_renderer = IconRenderer(cache_dir=str(self.cache_dir))

    def stat_box(
        self,
        number: str,
        label: str,
        sub_label: str = "",
        bg_color: str = "#1E40AF",
        text_color: str = "#FFFFFF",
        width: int = 320,
        height: int = 160,
        icon_name: Optional[str] = None,
    ) -> str:
        """Render a stat box: colored rounded rectangle with large number, label, and optional icon.

        Args:
            number: The main number/metric to display (large, bold).
            label: Label below the number.
            sub_label: Optional sub-label below the label.
            bg_color: Background color (hex).
            text_color: Text color (hex).
            width: Box width in pixels.
            height: Box height in pixels.
            icon_name: Optional icon to show in the top-right corner.

        Returns:
            Path to the rendered PNG file.
        """
        key = _cache_key(f"statbox_{number}_{label}_{bg_color}", width, text_color, bg_color, "box")
        cache_path = self.cache_dir / f"{key}.png"

        if cache_path.exists():
            return str(cache_path)

        img = Image.new("RGBA", (width, height), (0, 0, 0, 0))
        draw = ImageDraw.Draw(img)

        # Draw rounded rectangle background
        r, g, b = self._hex_to_rgb(bg_color)
        margin = 8
        radius = 16
        draw.rounded_rectangle(
            [margin, margin, width - margin, height - margin],
            radius=radius,
            fill=(r, g, b, 255),
        )

        # Draw number (large, centered at top)
        try:
            font_large = self._get_font(size=42, bold=True)
        except Exception:
            font_large = ImageFont.load_default()

        tr, tg, tb = self._hex_to_rgb(text_color)
        num_bbox = draw.textbbox((0, 0), number, font=font_large)
        num_w = num_bbox[2] - num_bbox[0]
        num_x = (width - num_w) // 2 - num_bbox[0]
        num_y = 20
        draw.text((num_x, num_y), number, fill=(tr, tg, tb, 255), font=font_large)

        # Draw label
        try:
            font_label = self._get_font(size=18, bold=False)
        except Exception:
            font_label = ImageFont.load_default()

        lbl_bbox = draw.textbbox((0, 0), label, font=font_label)
        lbl_w = lbl_bbox[2] - lbl_bbox[0]
        lbl_x = (width - lbl_w) // 2 - lbl_bbox[0]
        lbl_y = num_y + (num_bbox[3] - num_bbox[1]) + 10
        draw.text((lbl_x, lbl_y), label, fill=(tr, tg, tb, 255), font=font_label)

        # Draw sub-label
        if sub_label:
            try:
                font_sub = self._get_font(size=13, bold=False)
            except Exception:
                font_sub = ImageFont.load_default()

            # Slightly transparent
            sub_color = (tr, tg, tb, 180)
            sub_bbox = draw.textbbox((0, 0), sub_label, font=font_sub)
            sub_w = sub_bbox[2] - sub_bbox[0]
            sub_x = (width - sub_w) // 2 - sub_bbox[0]
            sub_y = lbl_y + (lbl_bbox[3] - lbl_bbox[1]) + 6
            draw.text((sub_x, sub_y), sub_label, fill=sub_color, font=font_sub)

        # Draw icon in top-right corner
        if icon_name:
            icon_path = self.icon_renderer.render(
                icon_name=icon_name,
                size=32,
                color=text_color,
                bg_color=bg_color,
                bg_shape="circle",
            )
            try:
                icon_img = Image.open(icon_path).convert("RGBA")
                icon_img = icon_img.resize((32, 32))
                img.paste(icon_img, (width - 44, 10), icon_img)
            except Exception as e:
                log.warning(f"Failed to place icon on stat box: {e}")

        cache_path.parent.mkdir(parents=True, exist_ok=True)
        img.save(str(cache_path), "PNG")
        return str(cache_path)

    def stat_boxes_row(
        self,
        boxes: list[dict],
        theme: Optional[Theme] = None,
        box_width: int = 320,
        box_height: int = 160,
        gap: int = 20,
    ) -> str:
        """Render a row of stat boxes side by side.

        Args:
            boxes: List of dicts with keys: number, label, sub_label, icon_name (optional).
                   Each dict can also override bg_color and text_color.
            theme: Theme for default colors.
            box_width: Width of each box.
            box_height: Height of each box.
            gap: Gap between boxes.

        Returns:
            Path to the rendered PNG file.
        """
        if not theme:
            theme = get_theme("corporate")
        c = theme.colors

        total_width = len(boxes) * box_width + (len(boxes) - 1) * gap
        img = Image.new("RGBA", (total_width, box_height), (0, 0, 0, 0))

        for i, box_data in enumerate(boxes):
            box_path = self.stat_box(
                number=str(box_data.get("number", "0")),
                label=str(box_data.get("label", "")),
                sub_label=str(box_data.get("sub_label", "")),
                bg_color=box_data.get("bg_color", c.primary),
                text_color=box_data.get("text_color", c.header_text),
                width=box_width,
                height=box_height,
                icon_name=box_data.get("icon_name"),
            )
            try:
                box_img = Image.open(box_path).convert("RGBA")
                x = i * (box_width + gap)
                img.paste(box_img, (x, 0), box_img)
            except Exception as e:
                log.warning(f"Failed to composite stat box {i}: {e}")

        key = _cache_key(f"statrow_{len(boxes)}", box_width, box_height, gap, "row")
        cache_path = self.cache_dir / f"{key}.png"
        cache_path.parent.mkdir(parents=True, exist_ok=True)
        img.save(str(cache_path), "PNG")
        return str(cache_path)

    def numbered_circle(
        self,
        number: int,
        size: int = 64,
        bg_color: str = "#1E40AF",
        text_color: str = "#FFFFFF",
    ) -> str:
        """Delegate to IconRenderer for numbered circles."""
        return self.icon_renderer.render_numbered_circle(
            number=number, size=size, bg_color=bg_color, text_color=text_color
        )

    def agenda_card(
        self,
        number: int,
        title: str,
        subtitle: str = "",
        theme: Optional[Theme] = None,
        card_width: int = 500,
        card_height: int = 120,
    ) -> str:
        """Render an agenda card with numbered circle, title, and subtitle.

        Args:
            number: Agenda item number.
            title: Agenda item title.
            subtitle: Optional subtitle.
            theme: Theme for colors.
            card_width: Card width in pixels.
            card_height: Card height in pixels.

        Returns:
            Path to the rendered PNG file.
        """
        if not theme:
            theme = get_theme("corporate")
        c = theme.colors

        key = _cache_key(f"agenda_{number}_{title}", card_width, card_height, str(c.primary), "card")
        cache_path = self.cache_dir / f"{key}.png"

        if cache_path.exists():
            return str(cache_path)

        img = Image.new("RGBA", (card_width, card_height), (0, 0, 0, 0))
        draw = ImageDraw.Draw(img)

        # Draw card background (light, with rounded corners)
        r, g, b = self._hex_to_rgb(c.table_alt_row)
        margin = 4
        radius = 12
        draw.rounded_rectangle(
            [margin, margin, card_width - margin, card_height - margin],
            radius=radius,
            fill=(r, g, b, 255),
        )

        # Draw numbered circle
        num_size = 48
        num_path = self.icon_renderer.render_numbered_circle(
            number=number,
            size=num_size,
            bg_color=c.primary,
            text_color=c.header_text,
        )
        try:
            num_img = Image.open(num_path).convert("RGBA")
            num_img = num_img.resize((num_size, num_size))
            img.paste(num_img, (20, (card_height - num_size) // 2), num_img)
        except Exception as e:
            log.warning(f"Failed to place number on agenda card: {e}")

        # Draw title
        try:
            font_title = self._get_font(size=20, bold=True)
        except Exception:
            font_title = ImageFont.load_default()

        tr, tg, tb = self._hex_to_rgb(c.text)
        title_x = num_size + 20
        title_y = (card_height - 40) // 2 - 12
        draw.text((title_x, title_y), title, fill=(tr, tg, tb, 255), font=font_title)

        # Draw subtitle
        if subtitle:
            try:
                font_sub = self._get_font(size=13, bold=False)
            except Exception:
                font_sub = ImageFont.load_default()

            slr, slg, slb = self._hex_to_rgb(c.text_light)
            sub_color = (slr, slg, slb, 200)
            title_bbox = draw.textbbox((0, 0), title, font=font_title)
            title_h = title_bbox[3] - title_bbox[1]
            sub_y = title_y + title_h + 4
            draw.text((title_x, sub_y), subtitle, fill=sub_color, font=font_sub)

        cache_path.parent.mkdir(parents=True, exist_ok=True)
        img.save(str(cache_path), "PNG")
        return str(cache_path)

    def timeline_item(
        self,
        step: int,
        title: str,
        date: str = "",
        description: str = "",
        theme: Optional[Theme] = None,
        item_width: int = 400,
        item_height: int = 100,
    ) -> str:
        """Render a timeline item with step number, title, date, and description.

        Args:
            step: Step number.
            title: Event title.
            date: Event date.
            description: Optional description.
            theme: Theme for colors.
            item_width: Item width in pixels.
            item_height: Item height in pixels.

        Returns:
            Path to the rendered PNG file.
        """
        if not theme:
            theme = get_theme("corporate")
        c = theme.colors

        key = _cache_key(f"timeline_{step}_{title}", item_width, item_height, str(c.primary), "item")
        cache_path = self.cache_dir / f"{key}.png"

        if cache_path.exists():
            return str(cache_path)

        img = Image.new("RGBA", (item_width, item_height), (0, 0, 0, 0))
        draw = ImageDraw.Draw(img)

        # Left border line
        br, bg, bb = self._hex_to_rgb(c.primary)
        draw.line([(10, 0), (10, item_height)], fill=(br, bg, bb, 255), width=3)

        # Circle on the line
        num_path = self.icon_renderer.render_numbered_circle(
            number=step,
            size=24,
            bg_color=c.primary,
            text_color=c.header_text,
        )
        try:
            num_img = Image.open(num_path).convert("RGBA")
            img.paste(num_img, (0, 8), num_img)
        except Exception as e:
            log.warning(f"Failed to place number on timeline: {e}")

        # Title
        try:
            font_title = self._get_font(size=16, bold=True)
        except Exception:
            font_title = ImageFont.load_default()

        tr, tg, tb = self._hex_to_rgb(c.text)
        draw.text((40, 10), title, fill=(tr, tg, tb, 255), font=font_title)

        # Date
        if date:
            try:
                font_date = self._get_font(size=12, bold=False)
            except Exception:
                font_date = ImageFont.load_default()

            dr, dg, db = self._hex_to_rgb(c.accent)
            draw.text((40, 35), date, fill=(dr, dg, db, 255), font=font_date)

        # Description
        if description:
            try:
                font_desc = self._get_font(size=11, bold=False)
            except Exception:
                font_desc = ImageFont.load_default()

            dlr, dlg, dlb = self._hex_to_rgb(c.text_light)
            draw.text((40, 58), description, fill=(dlr, dlg, dlb, 255), font=font_desc)

        cache_path.parent.mkdir(parents=True, exist_ok=True)
        img.save(str(cache_path), "PNG")
        return str(cache_path)

    def flow_diagram(
        self,
        nodes: list[dict],
        connections: list[tuple[int, int]],
        theme: Optional[Theme] = None,
        node_width: int = 140,
        node_height: int = 60,
        h_gap: int = 40,
        v_gap: int = 60,
    ) -> str:
        """Render a simple flow diagram with nodes and connecting arrows.

        Args:
            nodes: List of dicts with keys: label, sub_label (optional), icon_name (optional).
            connections: List of (from_index, to_index) tuples.
            theme: Theme for colors.
            node_width: Width of each node box.
            node_height: Height of each node box.
            h_gap: Horizontal gap between nodes.
            v_gap: Vertical gap between nodes.

        Returns:
            Path to the rendered PNG file.
        """
        if not theme:
            theme = get_theme("corporate")
        c = theme.colors

        # Calculate layout: arrange nodes in a grid
        cols = min(len(nodes), 4)
        rows = (len(nodes) + cols - 1) // cols

        total_w = cols * (node_width + h_gap) + h_gap
        total_h = rows * (node_height + v_gap) + v_gap

        img = Image.new("RGBA", (total_w, total_h), (0, 0, 0, 0))
        draw = ImageDraw.Draw(img)

        node_positions = []
        for i, node_data in enumerate(nodes):
            col = i % cols
            row = i // cols
            x = col * (node_width + h_gap) + h_gap
            y = row * (node_height + v_gap) + v_gap
            node_positions.append((x, y))

            # Draw node box
            nr, ng, nb = self._hex_to_rgb(c.primary)
            draw.rounded_rectangle(
                [x, y, x + node_width, y + node_height],
                radius=8,
                fill=(nr, ng, nb, 255),
            )

            # Draw node label
            try:
                font_label = self._get_font(size=14, bold=True)
            except Exception:
                font_label = ImageFont.load_default()

            label = str(node_data.get("label", f"Node {i + 1}"))
            tr, tg, tb = self._hex_to_rgb(c.header_text)
            lbl_bbox = draw.textbbox((0, 0), label, font=font_label)
            lbl_w = lbl_bbox[2] - lbl_bbox[0]
            lbl_x = x + (node_width - lbl_w) // 2 - lbl_bbox[0]
            lbl_y = y + 10
            draw.text((lbl_x, lbl_y), label, fill=(tr, tg, tb, 255), font=font_label)

            # Draw sub-label
            sub_label = node_data.get("sub_label", "")
            if sub_label:
                try:
                    font_sub = self._get_font(size=10, bold=False)
                except Exception:
                    font_sub = ImageFont.load_default()

                sub_color = (tr, tg, tb, 180)
                sub_bbox = draw.textbbox((0, 0), sub_label, font=font_sub)
                sub_w = sub_bbox[2] - sub_bbox[0]
                sub_x = x + (node_width - sub_w) // 2 - sub_bbox[0]
                sub_y = lbl_y + (lbl_bbox[3] - lbl_bbox[1]) + 4
                draw.text((sub_x, sub_y), sub_label, fill=sub_color, font=font_sub)

        # Draw connection arrows
        for from_idx, to_idx in connections:
            if 0 <= from_idx < len(node_positions) and 0 <= to_idx < len(node_positions):
                fx, fy = node_positions[from_idx]
                tx, ty = node_positions[to_idx]

                # Arrow from right of source to left of target (horizontal)
                # or bottom of source to top of target (vertical)
                if ty > fy:
                    # Vertical connection
                    start_y = fy + node_height
                    end_y = ty
                    cx = fx + node_width // 2
                    draw.line([(cx, start_y), (cx, end_y)], fill=(nr, ng, nb, 200), width=2)
                    # Arrowhead
                    draw.polygon([
                        (cx, end_y),
                        (cx - 6, end_y - 10),
                        (cx + 6, end_y - 10),
                    ], fill=(nr, ng, nb, 200))
                else:
                    # Horizontal connection
                    start_x = fx + node_width
                    end_x = tx
                    cy = fy + node_height // 2
                    draw.line([(start_x, cy), (end_x, cy)], fill=(nr, ng, nb, 200), width=2)
                    draw.polygon([
                        (end_x, cy),
                        (end_x - 10, cy - 6),
                        (end_x - 10, cy + 6),
                    ], fill=(nr, ng, nb, 200))

        key = _cache_key(f"flow_{len(nodes)}", total_w, total_h, str(c.primary), "diagram")
        cache_path = self.cache_dir / f"{key}.png"
        cache_path.parent.mkdir(parents=True, exist_ok=True)
        img.save(str(cache_path), "PNG")
        return str(cache_path)

    def comparison_card(
        self,
        title_left: str,
        title_right: str,
        items_left: list[str],
        items_right: list[str],
        theme: Optional[Theme] = None,
        width: int = 800,
        item_height: int = 30,
    ) -> str:
        """Render a side-by-side comparison card.

        Args:
            title_left: Left column title.
            title_right: Right column title.
            items_left: List of items for left column.
            items_right: List of items for right column.
            theme: Theme for colors.
            width: Total card width.
            item_height: Height per item row.

        Returns:
            Path to the rendered PNG file.
        """
        if not theme:
            theme = get_theme("corporate")
        c = theme.colors

        max_items = max(len(items_left), len(items_right))
        header_h = 50
        padding = 20
        total_h = header_h + max_items * item_height + padding * 2 + 20

        img = Image.new("RGBA", (width, total_h), (0, 0, 0, 0))
        draw = ImageDraw.Draw(img)

        # Background
        r, g, b = self._hex_to_rgb(c.background)
        draw.rounded_rectangle(
            [0, 0, width - 1, total_h - 1],
            radius=12,
            fill=(r, g, b, 255),
        )

        # Header bar
        hr, hg, hb = self._hex_to_rgb(c.primary)
        draw.rounded_rectangle(
            [0, 0, width - 1, header_h],
            radius=12,
            fill=(hr, hg, hb, 255),
        )

        # Headers text
        try:
            font_header = self._get_font(size=16, bold=True)
        except Exception:
            font_header = ImageFont.load_default()

        tr, tg, tb = self._hex_to_rgb(c.header_text)
        half_w = width // 2
        draw.text((padding + 10, 14), title_left, fill=(tr, tg, tb, 255), font=font_header)
        draw.text((half_w + 10, 14), title_right, fill=(tr, tg, tb, 255), font=font_header)

        # Divider line
        dr, dg, db = self._hex_to_rgb(c.border)
        draw.line([(half_w, header_h), (half_w, total_h)], fill=(dr, dg, db, 255), width=1)

        # Items
        try:
            font_item = self._get_font(size=12, bold=False)
        except Exception:
            font_item = ImageFont.load_default()

        text_color = self._hex_to_rgb(c.text)
        for i, item in enumerate(items_left):
            y = header_h + padding + i * item_height
            draw.text((padding + 10, y), f"• {item}", fill=text_color, font=font_item)

        for i, item in enumerate(items_right):
            y = header_h + padding + i * item_height
            draw.text((half_w + 10, y), f"• {item}", fill=text_color, font=font_item)

        key = _cache_key(f"compare_{title_left}_{title_right}", width, total_h, str(c.primary), "card")
        cache_path = self.cache_dir / f"{key}.png"
        cache_path.parent.mkdir(parents=True, exist_ok=True)
        img.save(str(cache_path), "PNG")
        return str(cache_path)

    def cover_slide_bg(
        self,
        theme: Optional[Theme] = None,
        width: int = 1280,
        height: int = 720,
    ) -> str:
        """Render a cover slide background image.

        Args:
            theme: Theme for colors.
            width: Image width.
            height: Image height.

        Returns:
            Path to the rendered PNG file.
        """
        if not theme:
            theme = get_theme("dark")
        c = theme.colors

        img = Image.new("RGBA", (width, height), (0, 0, 0, 0))
        draw = ImageDraw.Draw(img)

        # Full background
        br, bg, bb = self._hex_to_rgb(c.header_bg)
        draw.rectangle([0, 0, width, height], fill=(br, bg, bb, 255))

        # Accent stripe at bottom
        ar, ag, ab = self._hex_to_rgb(c.accent)
        stripe_h = 8
        draw.rectangle([0, height - stripe_h, width, height], fill=(ar, ag, ab, 255))

        # Subtle gradient overlay (simplified as semi-transparent rectangles)
        for i in range(0, height, 40):
            alpha = max(0, 30 - i // 40)
            draw.rectangle([0, i, width, i + 40], fill=(0, 0, 0, alpha))

        key = _cache_key(f"cover_bg_{theme.name}", width, height, str(c.header_bg), "cover")
        cache_path = self.cache_dir / f"{key}.png"
        cache_path.parent.mkdir(parents=True, exist_ok=True)
        img.save(str(cache_path), "PNG")
        return str(cache_path)

    # ── Font Helpers ──

    @staticmethod
    def _get_font(size: int, bold: bool = False) -> ImageFont.FreeTypeFont:
        """Load a font file. Tries common paths."""
        candidates = [
            "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf" if bold
            else "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
            "/usr/share/fonts/TTF/DejaVuSans-Bold.ttf" if bold
            else "/usr/share/fonts/TTF/DejaVuSans.ttf",
            "/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf" if bold
            else "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf",
        ]
        for path in candidates:
            try:
                return ImageFont.truetype(path, size)
            except (OSError, IOError):
                continue
        # Fallback to default
        return ImageFont.load_default()

    @staticmethod
    def _hex_to_rgb(hex_color: str) -> tuple[int, int, int]:
        """Convert hex color string to RGB tuple."""
        hex_color = hex_color.lstrip("#")
        return (
            int(hex_color[0:2], 16),
            int(hex_color[2:4], 16),
            int(hex_color[4:6], 16),
        )