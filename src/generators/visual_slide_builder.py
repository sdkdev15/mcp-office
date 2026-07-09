"""Visual slide builder -- creates premium slides with auto-icons and visual elements.

Works alongside PPTXGenerator to build cover slides, agenda slides, exec-summary slides,
timeline slides, and flow diagram slides with visual elements from VisualElements.
"""

from __future__ import annotations

from typing import Optional

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

from src.styles.themes import get_theme, Theme
from src.utils.colors import hex_to_rgbcolor_tuple
from src.utils.logger import get_logger
from src.utils.visual_elements import VisualElements
from src.utils.icon_library import find_icon_by_keyword, get_icon
from src.utils.icon_renderer import IconRenderer

log = get_logger("visual_slide_builder")


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

    def build_cover_slide(
        self,
        pres: Presentation,
        title: str,
        subtitle: str = "",
        icon_name: Optional[str] = None,
        extra_line: str = "",
    ) -> None:
        """Build a cover slide: dark bg + auto-detected icon + title + subtitle.

        Args:
            pres: Presentation object.
            title: Main title text.
            subtitle: Subtitle text.
            icon_name: Icon name (auto-detected from title if None).
            extra_line: Optional extra line below subtitle.
        """
        c = self.theme.colors

        # Auto-detect icon
        if not icon_name:
            icon_name = find_icon_by_keyword(title) or find_icon_by_keyword(subtitle)

        # Create blank slide
        slide_layout = pres.slide_layouts[6]  # blank
        slide = pres.slides.add_slide(slide_layout)

        # Full dark background
        bg_rect = slide.shapes.add_shape(
            1,  # MSO_SHAPE.RECTANGLE
            0, 0,
            Inches(13.33), Inches(7.5),
        )
        bg_rect.fill.solid()
        bg_rect.fill.fore_color.rgb = RGBColor(*hex_to_rgbcolor_tuple(c.header_bg))
        bg_rect.line.fill.background()
        self._send_to_back(bg_rect._element, 0)

        # Accent stripe at bottom
        accent = slide.shapes.add_shape(
            1,  # RECTANGLE
            0, Inches(7.42),
            Inches(13.33), Inches(0.08),
        )
        accent.fill.solid()
        accent.fill.fore_color.rgb = RGBColor(*hex_to_rgbcolor_tuple(c.accent))
        accent.line.fill.background()
        self._send_to_back(accent._element, 1)

        # Icon (top-right or centered)
        if icon_name:
            icon_path = self.icon_renderer.render(
                icon_name=icon_name,
                size=80,
                color=c.accent,
                bg_color=c.header_bg,
                bg_shape="circle",
            )
            try:
                slide.shapes.add_picture(
                    icon_path,
                    Inches(10.5), Inches(0.5),
                    Inches(1.2), Inches(1.2),
                )
            except Exception as e:
                log.warning(f"Failed to add icon to cover: {e}")

        # Title text box
        title_box = slide.shapes.add_textbox(
            Inches(1), Inches(2),
            Inches(11), Inches(2),
        )
        tf = title_box.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = title
        p.alignment = PP_ALIGN.LEFT
        for run in p.runs:
            run.font.size = Pt(36)
            run.font.bold = True
            run.font.color.rgb = RGBColor(255, 255, 255)
            run.font.name = self.theme.fonts.heading

        # Subtitle
        if subtitle:
            sub_box = slide.shapes.add_textbox(
                Inches(1), Inches(4.2),
                Inches(10), Inches(1),
            )
            stf = sub_box.text_frame
            stf.word_wrap = True
            sp = stf.paragraphs[0]
            sp.text = subtitle
            sp.alignment = PP_ALIGN.LEFT
            for run in sp.runs:
                run.font.size = Pt(16)
                run.font.italic = True
                run.font.color.rgb = RGBColor(*hex_to_rgbcolor_tuple(c.text_light))
                run.font.name = self.theme.fonts.body

        # Extra line
        if extra_line:
            # Divider line
            divider = slide.shapes.add_shape(
                1,  # RECTANGLE
                Inches(1), Inches(5.2),
                Inches(4), Inches(0.03),
            )
            divider.fill.solid()
            divider.fill.fore_color.rgb = RGBColor(*hex_to_rgbcolor_tuple(c.accent))
            divider.line.fill.background()

            extra_box = slide.shapes.add_textbox(
                Inches(1), Inches(5.4),
                Inches(10), Inches(0.5),
            )
            etf = extra_box.text_frame
            ep = etf.paragraphs[0]
            ep.text = extra_line
            for run in ep.runs:
                run.font.size = Pt(12)
                run.font.color.rgb = RGBColor(*hex_to_rgbcolor_tuple(c.text_light))

    def build_agenda_slide(
        self,
        pres: Presentation,
        title: str = "Agenda",
        items: Optional[list[dict]] = None,
    ) -> None:
        """Build an agenda slide: numbered grid of agenda items.

        Args:
            pres: Presentation object.
            title: Slide title.
            items: List of dicts with keys: title, subtitle (optional), icon (optional).
        """
        c = self.theme.colors
        slide_layout = pres.slide_layouts[6]  # blank
        slide = pres.slides.add_slide(slide_layout)

        # Slide title
        title_box = slide.shapes.add_textbox(
            Inches(0.8), Inches(0.3),
            Inches(11), Inches(0.8),
        )
        tf = title_box.text_frame
        p = tf.paragraphs[0]
        p.text = title
        p.alignment = PP_ALIGN.LEFT
        for run in p.runs:
            run.font.size = Pt(32)
            run.font.bold = True
            run.font.color.rgb = RGBColor(*hex_to_rgbcolor_tuple(c.primary))
            run.font.name = self.theme.fonts.heading

        # Grid layout: 2 columns, 3 rows
        cols = 2
        rows = (len(items) + cols - 1) // cols
        card_w = 5.5
        card_h = 1.4
        gap_x = 0.8
        gap_y = 0.3
        start_x = 0.8
        start_y = 1.4

        for i, item in enumerate(items):
            col = i % cols
            row = i // cols
            x = start_x + col * (card_w + gap_x)
            y = start_y + row * (card_h + gap_y)

            # Card background
            card = slide.shapes.add_shape(
                18,  # MSO_SHAPE.ROUNDED_RECTANGLE
                Inches(x), Inches(y),
                Inches(card_w), Inches(card_h),
            )
            cr, cg, cb = hex_to_rgbcolor_tuple(c.table_alt_row)
            card.fill.solid()
            card.fill.fore_color.rgb = RGBColor(cr, cg, cb)
            card.line.color.rgb = RGBColor(*hex_to_rgbcolor_tuple(c.border))
            card.line.width = Pt(1)

            # Numbered circle
            num_path = self.icon_renderer.render_numbered_circle(
                number=i + 1,
                size=40,
                bg_color=c.primary,
                text_color=c.header_text,
            )
            try:
                slide.shapes.add_picture(
                    num_path,
                    Inches(x + 0.15),
                    Inches(y + 0.2),
                    Inches(0.5), Inches(0.5),
                )
            except Exception as e:
                log.warning(f"Failed to add number to card {i}: {e}")

            # Item title
            item_title = str(item.get("title", f"Item {i + 1}"))
            title_box = slide.shapes.add_textbox(
                Inches(x + 0.75), Inches(y + 0.15),
                Inches(card_w - 1), Inches(0.5),
            )
            itf = title_box.text_frame
            itf.word_wrap = True
            ip = itf.paragraphs[0]
            ip.text = item_title
            for run in ip.runs:
                run.font.size = Pt(16)
                run.font.bold = True
                run.font.color.rgb = RGBColor(*hex_to_rgbcolor_tuple(c.text))
                run.font.name = self.theme.fonts.heading

            # Item subtitle
            subtitle = item.get("subtitle", "")
            if subtitle:
                sub_box = slide.shapes.add_textbox(
                    Inches(x + 0.75), Inches(y + 0.7),
                    Inches(card_w - 1), Inches(0.4),
                )
                stf = sub_box.text_frame
                stf.word_wrap = True
                sp = stf.paragraphs[0]
                sp.text = str(subtitle)
                for run in sp.runs:
                    run.font.size = Pt(11)
                    run.font.color.rgb = RGBColor(*hex_to_rgbcolor_tuple(c.text_light))

    def build_exec_summary_slide(
        self,
        pres: Presentation,
        title: str = "Ringkasan Eksekutif",
        stat_boxes: Optional[list[dict]] = None,
        body_text: str = "",
    ) -> None:
        """Build an exec summary slide: stat boxes + body paragraph.

        Args:
            pres: Presentation object.
            title: Slide title.
            stat_boxes: List of dicts with keys: number, label, sub_label, icon (optional),
                        bg_color (optional - uses theme primary by default).
            body_text: Body paragraph text below the stat boxes.
        """
        c = self.theme.colors
        slide_layout = pres.slide_layouts[6]  # blank
        slide = pres.slides.add_slide(slide_layout)

        # Slide title
        title_box = slide.shapes.add_textbox(
            Inches(0.8), Inches(0.3),
            Inches(11), Inches(0.8),
        )
        tf = title_box.text_frame
        p = tf.paragraphs[0]
        p.text = title
        p.alignment = PP_ALIGN.LEFT
        for run in p.runs:
            run.font.size = Pt(32)
            run.font.bold = True
            run.font.color.rgb = RGBColor(*hex_to_rgbcolor_tuple(c.primary))
            run.font.name = self.theme.fonts.heading

        # Stat boxes as colored rectangles
        num_boxes = len(stat_boxes)
        if num_boxes == 0:
            return

        box_w = min(3.5, (11.7 - 0.8 - (num_boxes - 1) * 0.3) / num_boxes)
        box_h = 1.5
        total_width = num_boxes * box_w + (num_boxes - 1) * 0.3
        start_x = (13.33 - total_width) / 2
        start_y = 1.4

        for i, box_data in enumerate(stat_boxes):
            x = start_x + i * (box_w + 0.3)
            y = start_y

            # Box background
            box_color = box_data.get("bg_color", c.primary)
            box_rect = slide.shapes.add_shape(
                18,  # ROUNDED_RECTANGLE
                Inches(x), Inches(y),
                Inches(box_w), Inches(box_h),
            )
            br, bg, bb = hex_to_rgbcolor_tuple(box_color)
            box_rect.fill.solid()
            box_rect.fill.fore_color.rgb = RGBColor(br, bg, bb)
            box_rect.line.fill.background()

            # Number
            number = str(box_data.get("number", "0"))
            num_box = slide.shapes.add_textbox(
                Inches(x), Inches(y + 0.15),
                Inches(box_w), Inches(0.5),
            )
            ntf = num_box.text_frame
            np = ntf.paragraphs[0]
            np.text = number
            np.alignment = PP_ALIGN.CENTER
            for run in np.runs:
                run.font.size = Pt(36)
                run.font.bold = True
                run.font.color.rgb = RGBColor(*hex_to_rgbcolor_tuple(c.header_text))
                run.font.name = self.theme.fonts.heading

            # Label
            label = str(box_data.get("label", ""))
            if label:
                lbl_box = slide.shapes.add_textbox(
                    Inches(x), Inches(y + 0.7),
                    Inches(box_w), Inches(0.4),
                )
                ltf = lbl_box.text_frame
                ltf.word_wrap = True
                lp = ltf.paragraphs[0]
                lp.text = label
                lp.alignment = PP_ALIGN.CENTER
                for run in lp.runs:
                    run.font.size = Pt(12)
                    run.font.color.rgb = RGBColor(*hex_to_rgbcolor_tuple(c.header_text))

            # Sub-label
            sub_label = box_data.get("sub_label", "")
            if sub_label:
                sub_box = slide.shapes.add_textbox(
                    Inches(x), Inches(y + 1.05),
                    Inches(box_w), Inches(0.3),
                )
                stf = sub_box.text_frame
                stf.word_wrap = True
                sp = stf.paragraphs[0]
                sp.text = str(sub_label)
                sp.alignment = PP_ALIGN.CENTER
                for run in sp.runs:
                    run.font.size = Pt(10)
                    run.font.color.rgb = RGBColor(*hex_to_rgbcolor_tuple(c.header_text))

            # Icon on stat box
            icon_name = box_data.get("icon")
            if not icon_name:
                icon_name = find_icon_by_keyword(label) or find_icon_by_keyword(str(number))
            if icon_name:
                icon_path = self.icon_renderer.render(
                    icon_name=icon_name,
                    size=36,
                    color=c.header_text,
                    bg_color=box_color,
                    bg_shape="none",
                )
                try:
                    slide.shapes.add_picture(
                        icon_path,
                        Inches(x + box_w - 0.6),
                        Inches(y + 0.1),
                        Inches(0.4), Inches(0.4),
                    )
                except Exception as e:
                    log.warning(f"Failed to add icon to stat box {i}: {e}")

        # Body text
        if body_text:
            body_box = slide.shapes.add_textbox(
                Inches(1), Inches(start_y + box_h + 0.5),
                Inches(11.33), Inches(3.5),
            )
            btf = body_box.text_frame
            btf.word_wrap = True
            bp = btf.paragraphs[0]
            bp.text = body_text
            for run in bp.runs:
                run.font.size = Pt(13)
                run.font.color.rgb = RGBColor(*hex_to_rgbcolor_tuple(c.text))
                run.font.name = self.theme.fonts.body

    def build_timeline_slide(
        self,
        pres: Presentation,
        title: str = "Timeline",
        events: Optional[list[dict]] = None,
        horizontal: bool = True,
    ) -> None:
        """Build a timeline slide.

        Args:
            pres: Presentation object.
            title: Slide title.
            events: List of dicts with keys: title, date (optional), description (optional).
            horizontal: If True, arrange horizontally. If False, vertically.
        """
        c = self.theme.colors
        slide_layout = pres.slide_layouts[6]  # blank
        slide = pres.slides.add_slide(slide_layout)

        # Slide title
        title_box = slide.shapes.add_textbox(
            Inches(0.8), Inches(0.3),
            Inches(11), Inches(0.8),
        )
        tf = title_box.text_frame
        p = tf.paragraphs[0]
        p.text = title
        p.alignment = PP_ALIGN.LEFT
        for run in p.runs:
            run.font.size = Pt(32)
            run.font.bold = True
            run.font.color.rgb = RGBColor(*hex_to_rgbcolor_tuple(c.primary))
            run.font.name = self.theme.fonts.heading

        if horizontal:
            self._build_horizontal_timeline(slide, events, c)
        else:
            self._build_vertical_timeline(slide, events, c)

    def _build_horizontal_timeline(
        self,
        slide,
        events: list[dict],
        c: "ThemeColors",
    ) -> None:
        """Build horizontal timeline layout."""
        num_events = len(events)
        if num_events == 0:
            return

        # Timeline line
        start_x = 1.0
        end_x = 12.0
        line_y = 3.5

        line_shape = slide.shapes.add_shape(
            1,  # RECTANGLE
            Inches(start_x), Inches(line_y),
            Inches(end_x - start_x), Inches(0.04),
        )
        lr, lg, lb = hex_to_rgbcolor_tuple(c.primary)
        line_shape.fill.solid()
        line_shape.fill.fore_color.rgb = RGBColor(lr, lg, lb)
        line_shape.line.fill.background()

        # Event points
        spacing = (end_x - start_x) / max(1, num_events - 1) if num_events > 1 else 0
        center_x = (start_x + end_x) / 2 if num_events == 1 else 0

        for i, event in enumerate(events):
            if num_events == 1:
                x = center_x
            else:
                x = start_x + i * spacing

            # Circle marker
            num_path = self.icon_renderer.render_numbered_circle(
                number=i + 1,
                size=32,
                bg_color=c.primary,
                text_color=c.header_text,
            )
            try:
                slide.shapes.add_picture(
                    num_path,
                    Inches(x - 0.2),
                    Inches(line_y - 0.16),
                    Inches(0.35), Inches(0.35),
                )
            except Exception as e:
                log.warning(f"Failed to add timeline marker {i}: {e}")

            # Title (above line for even, below for odd)
            title_text = str(event.get("title", ""))
            date_text = event.get("date", "")
            desc_text = event.get("description", "")

            is_above = i % 2 == 0
            if is_above:
                text_y = line_y - 1.5
            else:
                text_y = line_y + 0.5

            text_box = slide.shapes.add_textbox(
                Inches(x - 1), Inches(text_y),
                Inches(2.2), Inches(1.2),
            )
            text_tf = text_box.text_frame
            text_tf.word_wrap = True

            # Title
            tp = text_tf.paragraphs[0]
            tp.text = title_text
            tp.alignment = PP_ALIGN.CENTER
            for run in tp.runs:
                run.font.size = Pt(12)
                run.font.bold = True
                run.font.color.rgb = RGBColor(*hex_to_rgbcolor_tuple(c.text))

            # Date
            if date_text:
                dp = text_tf.add_paragraph()
                dp.text = str(date_text)
                dp.alignment = PP_ALIGN.CENTER
                for run in dp.runs:
                    run.font.size = Pt(10)
                    run.font.color.rgb = RGBColor(*hex_to_rgbcolor_tuple(c.accent))

            # Description
            if desc_text:
                desc_p = text_tf.add_paragraph()
                desc_p.text = str(desc_text)
                desc_p.alignment = PP_ALIGN.CENTER
                for run in desc_p.runs:
                    run.font.size = Pt(9)
                    run.font.color.rgb = RGBColor(*hex_to_rgbcolor_tuple(c.text_light))

    def _build_vertical_timeline(
        self,
        slide,
        events: list[dict],
        c: "ThemeColors",
    ) -> None:
        """Build vertical timeline layout."""
        num_events = len(events)
        if num_events == 0:
            return

        line_x = 2.0
        start_y = 1.8
        end_y = 6.5
        line_len = end_y - start_y

        # Vertical line
        line_shape = slide.shapes.add_shape(
            1,  # RECTANGLE
            Inches(line_x), Inches(start_y),
            Inches(0.04), Inches(line_len),
        )
        lr, lg, lb = hex_to_rgbcolor_tuple(c.primary)
        line_shape.fill.solid()
        line_shape.fill.fore_color.rgb = RGBColor(lr, lg, lb)
        line_shape.line.fill.background()

        spacing = line_len / max(1, num_events - 1) if num_events > 1 else 0
        center_y = (start_y + end_y) / 2 if num_events == 1 else 0

        for i, event in enumerate(events):
            if num_events == 1:
                y = center_y
            else:
                y = start_y + i * spacing

            # Circle marker
            num_path = self.icon_renderer.render_numbered_circle(
                number=i + 1,
                size=28,
                bg_color=c.primary,
                text_color=c.header_text,
            )
            try:
                slide.shapes.add_picture(
                    num_path,
                    Inches(line_x - 0.15),
                    Inches(y - 0.15),
                    Inches(0.3), Inches(0.3),
                )
            except Exception as e:
                log.warning(f"Failed to add timeline marker {i}: {e}")

            # Event info to the right
            title_text = str(event.get("title", ""))
            date_text = event.get("date", "")

            text_box = slide.shapes.add_textbox(
                Inches(line_x + 0.5), Inches(y - 0.15),
                Inches(9), Inches(0.8),
            )
            text_tf = text_box.text_frame
            text_tf.word_wrap = True

            tp = text_tf.paragraphs[0]
            tp.text = title_text
            for run in tp.runs:
                run.font.size = Pt(14)
                run.font.bold = True
                run.font.color.rgb = RGBColor(*hex_to_rgbcolor_tuple(c.text))

            if date_text:
                dp = text_tf.add_paragraph()
                dp.text = str(date_text)
                for run in dp.runs:
                    run.font.size = Pt(11)
                    run.font.color.rgb = RGBColor(*hex_to_rgbcolor_tuple(c.accent))

    def build_flow_slide(
        self,
        pres: Presentation,
        title: str = "Architecture",
        nodes: Optional[list[dict]] = None,
        connections: Optional[list[tuple[int, int]]] = None,
    ) -> None:
        """Build a flow diagram slide.

        Args:
            pres: Presentation object.
            title: Slide title.
            nodes: List of dicts with keys: label, sub_label (optional), icon (optional).
            connections: List of (from_index, to_index) tuples.
        """
        c = self.theme.colors
        slide_layout = pres.slide_layouts[6]  # blank
        slide = pres.slides.add_slide(slide_layout)

        # Slide title
        title_box = slide.shapes.add_textbox(
            Inches(0.8), Inches(0.3),
            Inches(11), Inches(0.8),
        )
        tf = title_box.text_frame
        p = tf.paragraphs[0]
        p.text = title
        p.alignment = PP_ALIGN.LEFT
        for run in p.runs:
            run.font.size = Pt(32)
            run.font.bold = True
            run.font.color.rgb = RGBColor(*hex_to_rgbcolor_tuple(c.primary))
            run.font.name = self.theme.fonts.heading

        # Layout: arrange in a flow
        num_nodes = len(nodes)
        if num_nodes == 0:
            return

        node_w = 2.0
        node_h = 1.0
        start_x = 1.5
        start_y = 1.8

        # Simple: alternate positions for flow
        positions = []
        for i, node_data in enumerate(nodes):
            if num_nodes <= 3:
                # Horizontal
                x = start_x + i * (node_w + 0.8)
                y = start_y
            elif i < num_nodes // 2:
                # Top row
                x = start_x + i * (node_w + 0.8)
                y = start_y
            else:
                # Bottom row
                j = i - num_nodes // 2
                x = start_x + j * (node_w + 0.8)
                y = start_y + node_h + 1.2

            positions.append((x, y))

            # Node box
            box = slide.shapes.add_shape(
                18,  # ROUNDED_RECTANGLE
                Inches(x), Inches(y),
                Inches(node_w), Inches(node_h),
            )
            br, bg, bb = hex_to_rgbcolor_tuple(c.primary)
            box.fill.solid()
            box.fill.fore_color.rgb = RGBColor(br, bg, bb)
            box.line.fill.background()

            # Node label
            label = str(node_data.get("label", f"Node {i + 1}"))
            lbl_box = slide.shapes.add_textbox(
                Inches(x), Inches(y + 0.1),
                Inches(node_w), Inches(0.4),
            )
            ltf = lbl_box.text_frame
            ltf.word_wrap = True
            lp = ltf.paragraphs[0]
            lp.text = label
            lp.alignment = PP_ALIGN.CENTER
            for run in lp.runs:
                run.font.size = Pt(13)
                run.font.bold = True
                run.font.color.rgb = RGBColor(*hex_to_rgbcolor_tuple(c.header_text))
                run.font.name = self.theme.fonts.heading

            # Sub-label
            sub_label = node_data.get("sub_label", "")
            if sub_label:
                sub_box = slide.shapes.add_textbox(
                    Inches(x), Inches(y + 0.5),
                    Inches(node_w), Inches(0.4),
                )
                stf = sub_box.text_frame
                stf.word_wrap = True
                sp = stf.paragraphs[0]
                sp.text = str(sub_label)
                sp.alignment = PP_ALIGN.CENTER
                for run in sp.runs:
                    run.font.size = Pt(10)
                    run.font.color.rgb = RGBColor(*hex_to_rgbcolor_tuple(c.header_text))

            # Icon
            icon_name = node_data.get("icon")
            if not icon_name:
                icon_name = find_icon_by_keyword(label)
            if icon_name:
                icon_path = self.icon_renderer.render(
                    icon_name=icon_name,
                    size=32,
                    color=c.header_text,
                    bg_color=c.primary,
                    bg_shape="none",
                )
                try:
                    slide.shapes.add_picture(
                        icon_path,
                        Inches(x + node_w - 0.5),
                        Inches(y + 0.05),
                        Inches(0.35), Inches(0.35),
                    )
                except Exception as e:
                    log.warning(f"Failed to add icon to flow node {i}: {e}")

        # Draw connections
        for from_idx, to_idx in connections:
            if 0 <= from_idx < len(positions) and 0 <= to_idx < len(positions):
                fx, fy = positions[from_idx]
                tx, ty = positions[to_idx]

                # Arrow line
                end_x_pos = fx + node_w
                mid_y = fy + node_h / 2
                target_x = tx
                target_y = ty + node_h / 2

                if ty > fy:
                    # Vertical arrow
                    arrow_end_x = target_x + node_w / 2
                    arrow_start_y = fy + node_h
                    arrow_end_y = target_y
                    line = slide.shapes.add_shape(
                        1,  # RECTANGLE
                        Inches(arrow_end_x - 0.015), Inches(arrow_start_y),
                        Inches(0.03), Inches(arrow_end_y - arrow_start_y),
                    )
                    lr, lg, lb = hex_to_rgbcolor_tuple(c.primary)
                    line.fill.solid()
                    line.fill.fore_color.rgb = RGBColor(lr, lg, lb)
                    line.line.fill.background()
                else:
                    # Horizontal arrow
                    arrow = slide.shapes.add_connector(
                        1,  # connector type
                        Inches(end_x_pos), Inches(mid_y),
                        Inches(target_x), Inches(target_y),
                    )
                    arrow.line.color.rgb = RGBColor(*hex_to_rgbcolor_tuple(c.primary))
                    arrow.line.width = Pt(2)

    def _send_to_back(self, element, index: int = 0) -> None:
        """Send a shape element to the back (z-order)."""
        element.getparent().insert(2 + index, element)