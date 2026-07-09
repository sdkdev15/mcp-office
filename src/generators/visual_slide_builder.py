"""Visual slide builder -- creates premium slides with auto-icons and visual elements.

Works alongside PPTXGenerator to build cover slides, agenda slides, exec-summary slides,
timeline slides, and flow diagram slides with premium layouts.
"""

from __future__ import annotations

from typing import Optional

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

from src.styles.themes import get_theme, Theme
from src.utils.colors import hex_to_rgbcolor_tuple
from src.utils.logger import get_logger
from src.utils.visual_elements import VisualElements
from src.utils.icon_library import find_icon_by_keyword, get_icon
from src.utils.icon_renderer import IconRenderer

log = get_logger("visual_slide_builder")

# ── Constants ──
SLIDE_W = 13.33
SLIDE_H = 7.5
CARD_CORNER = 12  # Pt


class VisualSlideBuilder:
    """Build premium slides with auto-icons and visual elements.

    Usage:
        builder = VisualSlideBuilder(theme_name="corporate")
        builder.build_cover_slide(pres, title="...", subtitle="...")
    """

    def __init__(self, theme_name: str = "corporate"):
        self.theme_name = theme_name
        self.theme = get_theme(theme_name)
        self.visual = VisualElements()
        self.icon_renderer = IconRenderer()

    # ── Color helpers ──

    def _dark_bg(self) -> str:
        return self.theme.colors.header_bg

    def _accent(self) -> str:
        return self.theme.colors.accent

    def _text_on_dark(self) -> str:
        return self.theme.colors.header_text

    def _text_light_on_dark(self) -> str:
        return "#94A3B8"  # slate-400

    def _card_bg_dark(self) -> str:
        return "#1E293B"  # slightly lighter than bg

    def _text_on_light(self) -> str:
        return self.theme.colors.text

    def _text_light_on_light(self) -> str:
        return self.theme.colors.text_light

    # ── Shape helpers ──

    def _add_rect(
        self, slide, x: float, y: float, w: float, h: float,
        fill: str, send_back: int = 0,
    ):
        """Add a filled rectangle."""
        shape = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            Inches(x), Inches(y), Inches(w), Inches(h),
        )
        shape.fill.solid()
        shape.fill.fore_color.rgb = RGBColor(*hex_to_rgbcolor_tuple(fill))
        shape.line.fill.background()
        if send_back >= 0:
            self._send_to_back(shape._element, send_back)
        return shape

    def _add_rounded_rect(
        self, slide, x: float, y: float, w: float, h: float,
        fill: str, border: Optional[str] = None, border_width: int = 1,
    ):
        """Add a filled rounded rectangle."""
        shape = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE,
            Inches(x), Inches(y), Inches(w), Inches(h),
        )
        shape.fill.solid()
        shape.fill.fore_color.rgb = RGBColor(*hex_to_rgbcolor_tuple(fill))
        if border:
            shape.line.color.rgb = RGBColor(*hex_to_rgbcolor_tuple(border))
            shape.line.width = Pt(border_width)
        else:
            shape.line.fill.background()
        return shape

    def _add_text(
        self, slide, text: str,
        x: float, y: float, w: float, h: float,
        font_size: int = 14,
        bold: bool = False,
        color: str = "#FFFFFF",
        align: str = PP_ALIGN.LEFT,
        font_name: Optional[str] = None,
        italic: bool = False,
    ):
        """Add a text box with styling."""
        box = slide.shapes.add_textbox(
            Inches(x), Inches(y), Inches(w), Inches(h),
        )
        tf = box.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = text
        p.alignment = align
        for run in p.runs:
            run.font.size = Pt(font_size)
            run.font.bold = bold
            run.font.italic = italic
            run.font.color.rgb = RGBColor(*hex_to_rgbcolor_tuple(color))
            if font_name:
                run.font.name = font_name
        return box

    def _add_multiline_text(
        self, slide,
        lines: list[tuple[str, int, bool, str]],  # (text, size, bold, color)
        x: float, y: float, w: float, h: float,
        align: str = PP_ALIGN.LEFT,
        line_spacing: float = 1.3,
    ):
        """Add a text box with multiple styled lines."""
        box = slide.shapes.add_textbox(
            Inches(x), Inches(y), Inches(w), Inches(h),
        )
        tf = box.text_frame
        tf.word_wrap = True

        for idx, (text, size, bold, color) in enumerate(lines):
            if idx == 0:
                p = tf.paragraphs[0]
            else:
                p = tf.add_paragraph()
            p.text = text
            p.alignment = align
            p.space_after = Pt(int(size * line_spacing))
            for run in p.runs:
                run.font.size = Pt(size)
                run.font.bold = bold
                run.font.color.rgb = RGBColor(*hex_to_rgbcolor_tuple(color))
                run.font.name = self.theme.fonts.body
        return box

    def _add_icon_circle(
        self, slide, icon_name: str, number: Optional[int] = None,
        x: float = 0, y: float = 0, size: float = 1.0,
        icon_color: str = "#F59E0B", bg_color: str = "#0F172A",
    ):
        """Add an icon inside a colored circle."""
        if number is not None:
            icon_path = self.icon_renderer.render_numbered_circle(
                number=number,
                size=int(size * 96),
                bg_color=icon_color,
                text_color=bg_color,
            )
        elif icon_name:
            icon_path = self.icon_renderer.render(
                icon_name=icon_name,
                size=int(size * 96),
                color=icon_color,
                bg_color=bg_color,
                bg_shape="circle",
            )
        else:
            return
        try:
            slide.shapes.add_picture(
                icon_path,
                Inches(x), Inches(y),
                Inches(size), Inches(size),
            )
        except Exception as e:
            log.warning(f"Failed to add icon: {e}")

    def _add_arrow(
        self, slide, from_x: float, from_y: float, to_x: float, to_y: float,
        color: str = "#F59E0B",
    ):
        """Add a simple arrow connector between two points."""
        arrow = slide.shapes.add_connector(
            1,  # straight connector
            Inches(from_x), Inches(from_y),
            Inches(to_x), Inches(to_y),
        )
        arrow.line.color.rgb = RGBColor(*hex_to_rgbcolor_tuple(color))
        arrow.line.width = Pt(2)

    def _send_to_back(self, element, index: int = 0) -> None:
        """Send a shape element to the back (z-order)."""
        element.getparent().insert(2 + index, element)

    # ── Cover Slide ──

    def build_cover_slide(
        self,
        pres: Presentation,
        title: str,
        subtitle: str = "",
        icon_name: Optional[str] = None,
        extra_line: str = "",
    ) -> None:
        """Build a cover slide: dark bg + auto-detected icon + title + subtitle."""
        c = self.theme.colors

        if not icon_name:
            icon_name = find_icon_by_keyword(title) or find_icon_by_keyword(subtitle)

        slide_layout = pres.slide_layouts[6]
        slide = pres.slides.add_slide(slide_layout)

        # Full dark background
        self._add_rect(slide, 0, 0, SLIDE_W, SLIDE_H, c.header_bg, send_back=0)

        # Accent stripe at bottom
        self._add_rect(slide, 0, SLIDE_H - 0.1, SLIDE_W, 0.1, c.accent, send_back=1)

        # Decorative left accent bar
        self._add_rect(slide, 0.6, 2.5, 0.08, 2.5, c.accent, send_back=2)

        # Icon (centered near title)
        if icon_name:
            icon_x = 9.5
            icon_y = 1.5
            icon_size = 1.5
            # Circle bg for icon
            self._add_rounded_rect(
                slide, icon_x - 0.1, icon_y - 0.1, icon_size + 0.2, icon_size + 0.2,
                fill=c.header_bg,  # transparent-looking
            )
            self._add_icon_circle(
                slide, icon_name,
                x=icon_x, y=icon_y,
                size=icon_size,
                icon_color=c.accent,
                bg_color=c.header_bg,
            )

        # Title — large, bold
        self._add_text(
            slide, title,
            x=1.0, y=2.2, w=8.5, h=2.5,
            font_size=42, bold=True,
            color="#FFFFFF",
            font_name=self.theme.fonts.heading,
        )

        # Subtitle
        if subtitle:
            self._add_text(
                slide, subtitle,
                x=1.0, y=4.5, w=10, h=1.0,
                font_size=18, italic=True,
                color=self._text_light_on_dark(),
                font_name=self.theme.fonts.body,
            )

        # Extra line
        if extra_line:
            self._add_rect(slide, 1.0, 5.6, 3, 0.04, c.accent)
            self._add_text(
                slide, extra_line,
                x=1.0, y=5.8, w=10, h=0.5,
                font_size=12,
                color=self._text_light_on_dark(),
            )

    # ── Agenda Slide (Premium) ──

    def build_agenda_slide(
        self,
        pres: Presentation,
        title: str = "Agenda",
        items: Optional[list[dict]] = None,
    ) -> None:
        """Build an agenda slide with large numbered cards in a grid."""
        c = self.theme.colors
        slide_layout = pres.slide_layouts[6]
        slide = pres.slides.add_slide(slide_layout)

        # Header bar (full-width)
        self._add_rect(slide, 0, 0, SLIDE_W, 1.15, c.header_bg, send_back=3)
        self._add_rect(slide, 0, 1.15, SLIDE_W, 0.06, c.accent, send_back=2)

        # Title on header
        self._add_text(
            slide, title,
            x=0.8, y=0.25, w=11, h=0.7,
            font_size=36, bold=True,
            color="#FFFFFF",
            font_name=self.theme.fonts.heading,
        )

        if not items:
            return

        # Grid: 2 columns, items // 2 rows
        num = len(items)
        cols = 2
        rows = (num + cols - 1) // cols
        card_w = 5.4
        card_h = (SLIDE_H - 2.0) / rows - 0.2  # fill available height
        gap_x = 0.6
        start_x = 0.7
        start_y = 1.5

        for i, item in enumerate(items):
            col = i % cols
            row = i // cols
            x = start_x + col * (card_w + gap_x)
            y = start_y + row * (card_h + 0.2)

            # Card background
            self._add_rounded_rect(
                slide, x, y, card_w, card_h,
                fill=c.table_alt_row,
                border=c.border,
            )

            # Numbered circle — large
            num_size = 0.7
            num_x = x + 0.3
            num_y = y + 0.3
            self._add_icon_circle(
                slide, "",
                number=i + 1,
                x=num_x, y=num_y,
                size=num_size,
                icon_color=c.accent,
                bg_color=c.header_bg,
            )

            # Item title — bold
            item_title = str(item.get("title", f"Item {i + 1}"))
            self._add_text(
                slide, item_title,
                x=x + 1.2, y=y + 0.25,
                w=card_w - 1.5, h=0.5,
                font_size=16, bold=True,
                color=c.text,
                font_name=self.theme.fonts.heading,
            )

            # Item subtitle
            subtitle = item.get("subtitle", "")
            if subtitle:
                self._add_text(
                    slide, str(subtitle),
                    x=x + 1.2, y=y + 0.7,
                    w=card_w - 1.5, h=0.6,
                    font_size=12,
                    color=self._text_light_on_light(),
                )

    # ── Exec Summary Slide (Premium) ──

    def build_exec_summary_slide(
        self,
        pres: Presentation,
        title: str = "Ringkasan Eksekutif",
        stat_boxes: Optional[list[dict]] = None,
        body_text: str = "",
    ) -> None:
        """Build an exec summary slide with full-height stat column cards.

        Layout mirrors 'Mirrorbox' style: dark bg, large rounded cards
        that fill most of the slide height, accent-colored headers.
        """
        c = self.theme.colors
        slide_layout = pres.slide_layouts[6]
        slide = pres.slides.add_slide(slide_layout)

        # Full dark background
        self._add_rect(slide, 0, 0, SLIDE_W, SLIDE_H, c.header_bg, send_back=3)

        # Accent top stripe
        self._add_rect(slide, 0, 0, SLIDE_W, 0.06, c.accent, send_back=2)

        # Title
        self._add_text(
            slide, title,
            x=0.8, y=0.3, w=11, h=0.8,
            font_size=38, bold=True,
            color="#FFFFFF",
            font_name=self.theme.fonts.heading,
        )

        if not stat_boxes:
            return

        num_boxes = len(stat_boxes)
        # Layout: columns of equal width, full-height cards
        gap = 0.3
        total_gap = (num_boxes - 1) * gap + 1.6  # left + right margin
        card_w = (SLIDE_W - total_gap) / num_boxes
        card_h = SLIDE_H - 3.2  # fill most of remaining height
        start_x = 0.8
        start_y = 1.3

        # Accent colors for variety
        accent_colors = [c.accent, c.success, c.error, c.primary]

        for i, box_data in enumerate(stat_boxes):
            x = start_x + i * (card_w + gap)
            y = start_y

            box_color = box_data.get("bg_color", accent_colors[i % len(accent_colors)])

            # Full-height card
            self._add_rounded_rect(
                slide, x, y, card_w, card_h,
                fill=self._card_bg_dark(),
                border=box_color,
                border_width=2,
            )

            # Accent bar at top of card
            self._add_rect(
                slide, x + 0.1, y + 0.05, card_w - 0.2, 0.06,
                box_color,
            )

            # Number — large and bold, centered
            number = str(box_data.get("number", "0"))
            self._add_text(
                slide, number,
                x=x, y=y + 0.3,
                w=card_w, h=0.8,
                font_size=44, bold=True,
                color="#FFFFFF",
                align=PP_ALIGN.CENTER,
                font_name=self.theme.fonts.heading,
            )

            # Label — below number
            label = str(box_data.get("label", ""))
            if label:
                self._add_text(
                    slide, label,
                    x=x, y=y + 1.2,
                    w=card_w, h=0.5,
                    font_size=16, bold=True,
                    color=box_color,
                    align=PP_ALIGN.CENTER,
                    font_name=self.theme.fonts.heading,
                )

            # Sub-label
            sub_label = box_data.get("sub_label", "")
            if sub_label:
                self._add_text(
                    slide, str(sub_label),
                    x=x, y=y + 1.75,
                    w=card_w, h=0.5,
                    font_size=12,
                    color=self._text_light_on_dark(),
                    align=PP_ALIGN.CENTER,
                )

            # Icon at bottom of card
            icon_name = box_data.get("icon")
            if not icon_name:
                icon_name = find_icon_by_keyword(label)
            if icon_name:
                icon_size = 0.7
                icon_x = x + (card_w - icon_size) / 2
                icon_y = y + card_h - icon_size - 0.3
                self._add_icon_circle(
                    slide, icon_name,
                    x=icon_x, y=icon_y,
                    size=icon_size,
                    icon_color=box_color,
                    bg_color=self._card_bg_dark(),
                )

        # Body text at bottom
        if body_text:
            self._add_text(
                slide, body_text,
                x=1.0, y=SLIDE_H - 1.2,
                w=SLIDE_W - 2, h=0.8,
                font_size=13,
                color=self._text_light_on_dark(),
            )

    # ── Timeline Slide (Premium) ──

    def build_timeline_slide(
        self,
        pres: Presentation,
        title: str = "Timeline",
        events: Optional[list[dict]] = None,
        horizontal: bool = True,
    ) -> None:
        """Build a timeline slide with dark bg and full-width card layout."""
        c = self.theme.colors
        slide_layout = pres.slide_layouts[6]
        slide = pres.slides.add_slide(slide_layout)

        # Full dark background
        self._add_rect(slide, 0, 0, SLIDE_W, SLIDE_H, c.header_bg, send_back=4)

        # Accent top stripe
        self._add_rect(slide, 0, 0, SLIDE_W, 0.06, c.accent, send_back=3)

        # Title
        self._add_text(
            slide, title,
            x=0.8, y=0.3, w=11, h=0.8,
            font_size=38, bold=True,
            color="#FFFFFF",
            font_name=self.theme.fonts.heading,
        )

        if not events:
            return

        if horizontal:
            self._build_premium_horizontal_timeline(slide, events, c)
        else:
            self._build_vertical_timeline(slide, events, c)

    def _build_premium_horizontal_timeline(
        self, slide, events: list[dict], c: "ThemeColors",
    ) -> None:
        """Premium horizontal timeline: full-height cards with numbers."""
        num = len(events)
        if num == 0:
            return

        gap = 0.25
        margin = 0.7
        arrows_w = (num - 1) * 0.35 if num > 1 else 0
        total_content_w = SLIDE_W - margin * 2 - arrows_w
        card_w = (total_content_w - (num - 1) * gap) / num
        card_h = SLIDE_H - 3.0
        start_x = margin
        start_y = 1.4

        for i, event in enumerate(events):
            x = start_x + i * (card_w + gap + (0.35 if i < num - 1 else 0))
            y = start_y

            # Card background
            self._add_rounded_rect(
                slide, x, y, card_w, card_h,
                fill=self._card_bg_dark(),
            )

            # Numbered circle — large, top center
            num_size = 0.8
            num_x = x + (card_w - num_size) / 2
            num_y = y + 0.3
            self._add_icon_circle(
                slide, "",
                number=i + 1,
                x=num_x, y=num_y,
                size=num_size,
                icon_color=c.accent,
                bg_color=self._card_bg_dark(),
            )

            # Title
            title_text = str(event.get("title", ""))
            self._add_text(
                slide, title_text,
                x=x, y=y + 1.3,
                w=card_w, h=0.7,
                font_size=16, bold=True,
                color="#FFFFFF",
                align=PP_ALIGN.CENTER,
                font_name=self.theme.fonts.heading,
            )

            # Date
            date_text = event.get("date", "")
            if date_text:
                self._add_text(
                    slide, str(date_text),
                    x=x, y=y + 2.0,
                    w=card_w, h=0.4,
                    font_size=12,
                    color=c.accent,
                    align=PP_ALIGN.CENTER,
                )

            # Description
            desc_text = event.get("description", "")
            if desc_text:
                self._add_text(
                    slide, str(desc_text),
                    x=x + 0.2, y=y + 2.5,
                    w=card_w - 0.4, h=card_h - 3.0,
                    font_size=11,
                    color=self._text_light_on_dark(),
                    align=PP_ALIGN.CENTER,
                )

            # Arrow to next card
            if i < num - 1:
                arrow_x = x + card_w
                arrow_y = y + card_h / 2
                self._add_arrow(
                    slide, arrow_x, arrow_y,
                    arrow_x + 0.35, arrow_y,
                    color=c.accent,
                )

    # ── Flow Slide (Premium) ──

    def build_flow_slide(
        self,
        pres: Presentation,
        title: str = "Architecture",
        nodes: Optional[list[dict]] = None,
        connections: Optional[list[tuple[int, int]]] = None,
        columns: Optional[int] = None,
    ) -> None:
        """Build a flow diagram slide with configurable column layout.

        Args:
            pres: Presentation object.
            title: Slide title.
            nodes: List of dicts with keys: label, sub_label (optional), icon (optional).
            connections: List of (from_index, to_index) tuples.
            columns: Number of columns. Auto-calculated if None.
        """
        c = self.theme.colors
        slide_layout = pres.slide_layouts[6]
        slide = pres.slides.add_slide(slide_layout)

        # Header bar
        self._add_rect(slide, 0, 0, SLIDE_W, 1.15, c.header_bg, send_back=4)
        self._add_rect(slide, 0, 1.15, SLIDE_W, 0.06, c.accent, send_back=3)

        # Title
        self._add_text(
            slide, title,
            x=0.8, y=0.25, w=11, h=0.7,
            font_size=36, bold=True,
            color="#FFFFFF",
            font_name=self.theme.fonts.heading,
        )

        if not nodes:
            return

        num = len(nodes)
        if columns is None:
            columns = min(num, 5)
        rows = (num + columns - 1) // columns

        # Calculate card sizes to fill slide
        margin_x = 0.6
        margin_y = 1.5
        gap_x = 0.3
        gap_y = 0.5
        arrow_w = 0.4

        avail_w = SLIDE_W - margin_x * 2 - (columns - 1) * gap_x - (columns - 1) * arrow_w
        card_w = avail_w / columns
        avail_h = SLIDE_H - margin_y - 1.0  # leave room for bottom
        if rows > 1:
            avail_h -= (rows - 1) * gap_y
        card_h = min(avail_h, 4.5)  # cap height
        row_y_start = margin_y + (avail_h - card_h) / 2

        positions = []
        for i, node_data in enumerate(nodes):
            col = i % columns
            row = i // columns
            x = margin_x + col * (card_w + gap_x + arrow_w)
            y = row_y_start + row * (card_h + gap_y)
            positions.append((x, y))

            # Card background
            self._add_rounded_rect(
                slide, x, y, card_w, card_h,
                fill=c.table_alt_row,
                border=c.border,
            )

            # Numbered circle
            num_size = 0.6
            num_x = x + (card_w - num_size) / 2
            num_y = y + 0.2
            self._add_icon_circle(
                slide, "",
                number=i + 1,
                x=num_x, y=num_y,
                size=num_size,
                icon_color=c.accent,
                bg_color=c.header_bg,
            )

            # Title — bold
            label = str(node_data.get("label", f"Step {i + 1}"))
            self._add_text(
                slide, label,
                x=x, y=y + 1.0,
                w=card_w, h=0.6,
                font_size=14, bold=True,
                color=c.text,
                align=PP_ALIGN.CENTER,
                font_name=self.theme.fonts.heading,
            )

            # Sub-label / description
            sub_label = node_data.get("sub_label", "")
            if sub_label:
                self._add_text(
                    slide, str(sub_label),
                    x=x + 0.2, y=y + 1.6,
                    w=card_w - 0.4, h=card_h - 2.0,
                    font_size=11,
                    color=self._text_light_on_light(),
                    align=PP_ALIGN.CENTER,
                )

            # Arrow to next in row
            if i < num - 1:
                next_col = (i + 1) % columns
                if next_col == 0 or i + 1 >= num:
                    break
                # Only draw arrow if next item is in the next column (same row)
                next_row = (i + 1) // columns
                if next_row == row:
                    arrow_x = x + card_w
                    arrow_y = y + card_h / 2
                    self._add_arrow(
                        slide, arrow_x, arrow_y,
                        arrow_x + arrow_w, arrow_y,
                        color=c.accent,
                    )

        # Draw explicit connections
        if connections:
            for from_idx, to_idx in connections:
                if 0 <= from_idx < len(positions) and 0 <= to_idx < len(positions):
                    fx, fy = positions[from_idx]
                    tx, ty = positions[to_idx]
                    self._add_arrow(
                        slide,
                        fx + card_w, fy + card_h / 2,
                        tx, ty + card_h / 2,
                        color=c.accent,
                    )

    # ── Legacy vertical timeline (kept for backward compat) ──

    def _build_vertical_timeline(
        self, slide, events: list[dict], c: "ThemeColors",
    ) -> None:
        """Build vertical timeline layout."""
        num = len(events)
        if num == 0:
            return

        line_x = 1.5
        start_y = 1.8
        end_y = SLIDE_H - 0.8
        line_len = end_y - start_y

        # Vertical line
        self._add_rect(slide, line_x, start_y, 0.05, line_len, c.primary, send_back=0)

        spacing = line_len / max(1, num - 1) if num > 1 else 0
        center_y = (start_y + end_y) / 2 if num == 1 else 0

        for i, event in enumerate(events):
            y = center_y if num == 1 else start_y + i * spacing

            # Circle marker
            self._add_icon_circle(
                slide, "",
                number=i + 1,
                x=line_x - 0.15, y=y - 0.15,
                size=0.35,
                icon_color=c.accent,
                bg_color=c.header_bg,
            )

            # Event info
            title_text = str(event.get("title", ""))
            self._add_text(
                slide, title_text,
                x=line_x + 0.6, y=y - 0.2,
                w=9, h=0.5,
                font_size=14, bold=True,
                color="#FFFFFF",
                font_name=self.theme.fonts.heading,
            )

            date_text = event.get("date", "")
            if date_text:
                self._add_text(
                    slide, str(date_text),
                    x=line_x + 0.6, y=y + 0.35,
                    w=9, h=0.4,
                    font_size=11,
                    color=c.accent,
                )