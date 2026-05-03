"""PowerPoint presentation generator using python-pptx with theme support."""

from __future__ import annotations

import io
import re
from typing import Any, Optional

from lxml import etree
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.chart import XL_CHART_TYPE
from pptx.oxml.ns import qn
from pptx.oxml.xmlchemy import OxmlElement

from src.styles.themes import get_theme
from src.utils.logger import get_logger
from src.utils.validators import validate_slide_layout, ValidationError

log = get_logger("pptx_generator")


def hex_to_rgbcolor(hex_color: str) -> RGBColor:
    """Convert hex color string to RGBColor for python-pptx.

    Args:
        hex_color: Hex color string (e.g., '1E40AF' or '#1E40AF').

    Returns:
        RGBColor instance.
    """
    c = hex_color.lstrip("#")
    r = int(c[0:2], 16)
    g = int(c[2:4], 16)
    b = int(c[4:6], 16)
    return RGBColor(r, g, b)


class PPTXGenerator:
    """Generates PowerPoint presentations with rich slides, charts, and styling."""

    SLIDE_WIDTH_STANDARD_INCHES = 10.0
    SLIDE_HEIGHT_STANDARD_INCHES = 7.5
    SLIDE_WIDTH_WIDESCREEN_INCHES = 13.33
    SLIDE_HEIGHT_WIDESCREEN_INCHES = 7.5

    def __init__(self, theme_name: str = "corporate"):
        self.theme_name = theme_name
        self.theme = get_theme(theme_name)
        self.pres: Optional[Presentation] = None

    def create_presentation(
        self,
        title: str = "Presentation",
        slide_size: str = "widescreen",
        metadata: Optional[dict] = None,
        slides: Optional[list[dict]] = None,
    ) -> bytes:
        """Create a new PowerPoint presentation.

        Args:
            title: Presentation title.
            slide_size: Slide size (widescreen for 16:9, standard for 4:3).
            metadata: Optional document metadata.
            slides: Optional list of slides to add.

        Returns:
            Presentation content as bytes.
        """
        self.pres = Presentation()

        # Set slide size
        if slide_size.lower() == "standard":
            self.pres.slide_width = Inches(self.SLIDE_WIDTH_STANDARD_INCHES)
            self.pres.slide_height = Inches(self.SLIDE_HEIGHT_STANDARD_INCHES)
        else:  # widescreen
            self.pres.slide_width = Inches(self.SLIDE_WIDTH_WIDESCREEN_INCHES)
            self.pres.slide_height = Inches(self.SLIDE_HEIGHT_WIDESCREEN_INCHES)

        # Apply metadata
        if metadata:
            self._apply_metadata(metadata)

        # Add title slide
        self.add_slide("title", title=title)

        # Add content slides
        if slides:
            for slide_data in slides:
                self.add_slide(
                    layout=slide_data.get("layout", "title_and_content"),
                    title=slide_data.get("title"),
                    content=slide_data.get("content"),
                    bullets=slide_data.get("bullets"),
                )

        return self._save_presentation()

    def add_slide(
        self,
        layout: str = "title_and_content",
        title: Optional[str] = None,
        content: Optional[str] = None,
        bullets: Optional[list[str]] = None,
    ) -> int:
        """Add a slide to the presentation.

        Args:
            layout: Slide layout type.
            title: Slide title.
            content: Body content text.
            bullets: List of bullet points.

        Returns:
            Index of the added slide.
        """
        layout = validate_slide_layout(layout)
        slide_layout_map = {
            "title": 0,
            "title_and_content": 1,
            "blank": 6,
            "two_content": 2,
            "comparison": 5,
            "title_only": 5,
            "section_header": 11,
        }

        layout_idx = slide_layout_map.get(layout, 1)

        try:
            slide_layout = self.pres.slide_layouts[layout_idx]
        except IndexError:
            slide_layout = self.pres.slide_layouts[1]

        slide = self.pres.slides.add_slide(slide_layout)
        slide_idx = len(self.pres.slides) - 1

        # Set title
        if title:
            try:
                if getattr(slide.shapes, 'title', None):
                    slide.shapes.title.text = title
                    self._style_title(slide.shapes.title)
            except Exception as e:
                log.warning(f"Failed to set title for slide '{title}': {e}")

        # Add content
        if content or bullets:
            try:
                body_shape = None
                for i, ph in enumerate(slide.placeholders):
                    if ph.placeholder_type.name == "BODY":
                        body_shape = slide.placeholders[i]
                        break
                if body_shape is None:
                    body_shape = slide.placeholders[1]

                tf = body_shape.text_frame

                # Remove existing paragraphs
                while tf.paragraphs:
                    existing_p = tf.paragraphs[0]
                    tf._remove(existing_p._p)

                if content:
                    p = tf.add_paragraph()
                    p.text = content
                    # Style the run that was created by setting p.text
                    for run in p.runs:
                        self._style_run(run)

                if bullets:
                    for idx, bullet in enumerate(bullets):
                        p = tf.add_paragraph()
                        p.text = bullet
                        p.level = 0
                        # Set bullet formatting via XML
                        pPr = p._p.get_or_add_pPr()
                        buChar = OxmlElement('a:buChar')
                        buChar.set('char', '\u2022')
                        pPr.append(buChar)
                        # Style the run that was created by setting p.text
                        for run in p.runs:
                            self._style_run(run)
            except Exception as e:
                log.warning(f"Failed to format body text: {e}")
                if content or bullets:
                    text = content or "\n".join(bullets) if bullets else ""
                    self.add_text_box(slide_idx, text, left=1, top=2)

        return slide_idx

    def add_text_box(
        self,
        slide_index: int,
        text: str,
        left: float = 1.0,
        top: float = 1.0,
        width: float = 5.0,
        height: float = 1.0,
        font_size: Optional[int] = None,
        bold: bool = False,
        color: Optional[str] = None,
    ) -> None:
        """Add a text box to a slide."""
        slide = self.pres.slides[slide_index]
        txBox = slide.shapes.add_textbox(
            Inches(left), Inches(top), Inches(width), Inches(height)
        )
        tf = txBox.text_frame
        p = tf.add_paragraph()
        p.text = text

        run = p.runs[0] if p.runs else p.add_run()
        run.bold = bold
        run.font.size = Pt(font_size or self.theme.fonts.body_size)

        if color:
            run.font.color.rgb = hex_to_rgbcolor(color)
        else:
            run.font.color.rgb = hex_to_rgbcolor(self.theme.colors.text)

    def add_image(
        self,
        slide_index: int,
        image_path: str,
        left: float = 1.0,
        top: float = 1.0,
        width: Optional[float] = None,
        height: Optional[float] = None,
    ) -> None:
        """Add an image to a slide."""
        slide = self.pres.slides[slide_index]
        slide.shapes.add_picture(
            image_path,
            Inches(left),
            Inches(top),
            width=Inches(width) if width else None,
            height=Inches(height) if height else None,
        )

    def add_table(
        self,
        slide_index: int,
        headers: list[str],
        rows: list[list[Any]],
    ) -> None:
        """Add a table to a slide."""
        slide = self.pres.slides[slide_index]
        num_rows = len(rows) + 1
        num_cols = len(headers)

        table_width = Inches(min(num_cols * 1.5, 10))
        table_height = Inches(min(num_rows * 0.4, 5))

        table_shape = slide.shapes.add_table(
            num_rows, num_cols, Inches(1), Inches(1.5), table_width, table_height
        )
        table = table_shape.table
        colors = self.theme.colors

        # Set headers
        for i, header in enumerate(headers):
            cell = table.cell(0, i)
            cell.text = str(header)
            for paragraph in cell.text_frame.paragraphs:
                paragraph.alignment = PP_ALIGN.CENTER
                for run in paragraph.runs:
                    run.bold = True
                    run.font.size = Pt(11)
                    run.font.color.rgb = hex_to_rgbcolor(colors.header_text)

        # Set data
        for row_idx, row_data in enumerate(rows):
            for col_idx, value in enumerate(row_data):
                cell = table.cell(row_idx + 1, col_idx)
                cell.text = str(value)
                for paragraph in cell.text_frame.paragraphs:
                    for run in paragraph.runs:
                        run.font.size = Pt(10)
                        run.font.color.rgb = hex_to_rgbcolor(colors.text)

    def add_chart(
        self,
        slide_index: int,
        chart_type: str,
        headers: list[str],
        rows: list[list[Any]],
        title: str = "Chart",
    ) -> None:
        """Add a chart to a slide."""
        chart_type_map = {
            "bar": XL_CHART_TYPE.BAR_CLUSTERED,
            "column": XL_CHART_TYPE.COLUMN_CLUSTERED,
            "line": XL_CHART_TYPE.LINE,
            "pie": XL_CHART_TYPE.PIE,
            "area": XL_CHART_TYPE.AREA,
            "doughnut": XL_CHART_TYPE.DOUGHNUT,
        }

        chart_type_enum = chart_type_map.get(chart_type.lower(), XL_CHART_TYPE.COLUMN_CLUSTERED)
        slide = self.pres.slides[slide_index]

        all_data = [headers] + [[str(cell) for cell in row] for row in rows]
        chart_data = self._create_chart_data(all_data)

        chart_frame = slide.shapes.add_chart(
            chart_type_enum, Inches(1), Inches(1), Inches(9), Inches(5), chart_data
        )
        chart = chart_frame.chart
        chart.has_title = True
        chart.chart_title.text_frame.text = title

    def set_background(self, slide_index: int, color: Optional[str] = None) -> None:
        """Set slide background color."""
        slide = self.pres.slides[slide_index]
        fill = slide.background.fill
        fill.solid()
        bg_color = color or self.theme.colors.background
        fill.fore_color.rgb = hex_to_rgbcolor(bg_color)

    def create_from_prompt(self, prompt: str, title: str = "Presentation") -> bytes:
        """Create a presentation from a natural language prompt.

        Parses the prompt to extract slides from "Heading - content" lines.
        """
        log.info(f"Creating presentation from prompt: {prompt[:100]}...")

        self.pres = Presentation()
        self.pres.slide_width = Inches(self.SLIDE_WIDTH_WIDESCREEN_INCHES)
        self.pres.slide_height = Inches(self.SLIDE_HEIGHT_WIDESCREEN_INCHES)

        # Extract title from prompt if generic
        doc_title = title if title != "Presentation" else self._extract_title(prompt)
        self.add_slide("title", title=doc_title)

        # Parse prompt into slides
        self._parse_prompt_to_slides(prompt)

        return self._save_presentation()

    def _extract_title(self, text: str) -> str:
        """Extract presentation title from prompt text."""
        lines = text.strip().split("\n")
        for line in lines:
            line = line.strip()
            if re.match(r"^(Buat|Format|Gunakan|Create|Generate)", line, re.IGNORECASE):
                continue
            if line and len(line) > 10:
                return line.rstrip(":").strip()
        return "Presentation"

    def _parse_prompt_to_slides(self, text: str) -> None:
        """Parse prompt text into slides."""
        lines = text.strip().split("\n")

        for line in lines:
            line = line.strip()
            if not line:
                continue

            # Skip instruction lines
            if re.match(r"^(Buat|Format|Gunakan|Create|Generate|Silakan|Please)", line, re.IGNORECASE):
                continue

            # Detect "Heading - content" pattern (lenient: first word capitalized, 2+ words before dash)
            dash_match = re.match(r"^([A-Z][^-\n]{5,}?)\s+-\s+(.+)$", line)
            if dash_match:
                slide_title = dash_match.group(1).strip()
                content_text = dash_match.group(2).strip()

                # Try to split content into bullets (comma-separated)
                if "," in content_text and len(content_text) > 50:
                    parts = [p.strip() for p in content_text.split(",") if p.strip()]
                    if len(parts) > 1:
                        self.add_slide(
                            title=slide_title,
                            bullets=parts,
                        )
                        continue

                self.add_slide(
                    title=slide_title,
                    content=content_text,
                )
                continue

            # ALL CAPS heading
            if line.isupper() and not line.startswith(("-", "*", "•")):
                self.add_slide(title=line.title())
                continue

            # Bullet points - add to last slide
            bullet_match = re.match(r"^[-*•]\s+(.*)", line)
            if bullet_match:
                # Would need to append to last slide, skip for now
                continue

            # Regular paragraph - treat as content slide
            if len(line) > 20:
                self.add_slide(title="Content", content=line)

    def create_from_template(
        self,
        template_path: str,
        slides: Optional[list[dict]] = None,
        metadata: Optional[dict] = None,
    ) -> bytes:
        """Create a presentation using an existing .pptx file as template.

        The template preserves all master slide designs, themes, fonts,
        and layouts. New slides are appended after existing content.

        Args:
            template_path: Path to the .pptx template file.
            slides: Optional list of slide dicts with title, content, bullets.
            metadata: Optional document metadata.

        Returns:
            Presentation content as bytes.
        """
        self.pres = Presentation(template_path)

        # Apply metadata
        if metadata:
            self._apply_metadata(metadata)

        # Add new slides
        if slides:
            for slide_data in slides:
                self.add_slide(
                    layout=slide_data.get("layout", "title_and_content"),
                    title=slide_data.get("title"),
                    content=slide_data.get("content"),
                    bullets=slide_data.get("bullets"),
                )

        return self._save_presentation()

    def _style_title(self, shape) -> None:
        """Apply theme styling to a title shape."""
        for paragraph in shape.text_frame.paragraphs:
            paragraph.alignment = PP_ALIGN.CENTER
            # If no runs exist, add one
            if not paragraph.runs:
                paragraph.add_run()
            for run in paragraph.runs:
                run.font.name = self.theme.fonts.heading
                run.font.size = Pt(self.theme.fonts.heading1_size)
                run.font.color.rgb = hex_to_rgbcolor(self.theme.colors.primary)

    def _style_run(self, run) -> None:
        """Apply theme styling to a single run."""
        run.font.name = self.theme.fonts.body
        run.font.size = Pt(self.theme.fonts.body_size)
        run.font.color.rgb = hex_to_rgbcolor(self.theme.colors.text)

    def _apply_metadata(self, metadata: dict) -> None:
        """Apply presentation metadata."""
        props = self.pres.core_properties
        if metadata.get("author"):
            props.author = metadata["author"]
        if metadata.get("company"):
            props.company = metadata["company"]
        if metadata.get("subject"):
            props.subject = metadata["subject"]
        if metadata.get("title"):
            props.title = metadata["title"]
        if metadata.get("keywords"):
            props.keywords = metadata["keywords"]
        if metadata.get("category"):
            props.category = metadata["category"]
        if metadata.get("comments"):
            props.description = metadata["comments"]

    def _create_chart_data(self, data: list[list]) -> object:
        """Create chart data from 2D array."""
        from pptx.chart.data import CategoryChartData

        chart_data = CategoryChartData()
        chart_data.categories = [str(item) for item in data[0][1:]] if len(data) > 1 else []

        for row in data[1:]:
            chart_data.add_series(str(row[0]), [float(x) if x else 0 for x in row[1:]])

        return chart_data

    def _save_presentation(self) -> bytes:
        """Save presentation to bytes."""
        buffer = io.BytesIO()
        self.pres.save(buffer)
        buffer.seek(0)
        return buffer.getvalue()