"""Pydantic models for tool input/output schemas."""

from __future__ import annotations

from typing import Any, Optional
from enum import Enum

from pydantic import BaseModel, Field


# ── Common Models ──

class ExportFormat(str, Enum):
    xlsx = "xlsx"
    ods = "ods"
    csv = "csv"
    pptx = "pptx"
    odp = "odp"
    docx = "docx"
    odt = "odt"
    pdf = "pdf"
    all = "all"


class ChartType(str, Enum):
    bar = "bar"
    column = "column"
    line = "line"
    pie = "pie"
    doughnut = "doughnut"
    area = "area"
    radar = "radar"
    scatter = "scatter"
    bubble = "bubble"


class SlideLayout(str, Enum):
    title = "title"
    title_and_content = "title_and_content"
    blank = "blank"
    two_content = "two_content"
    comparison = "comparison"
    section_header = "section_header"
    title_only = "title_only"


class DocumentMetadata(BaseModel):
    """Metadata for generated documents."""
    author: Optional[str] = None
    company: Optional[str] = None
    subject: Optional[str] = None
    title: Optional[str] = None
    keywords: Optional[str] = None
    category: Optional[str] = None
    comments: Optional[str] = None


class ExportRequest(BaseModel):
    """Request model for exporting documents in multiple formats."""
    filepath: str = Field(..., description="Path to the source file")
    formats: list[ExportFormat] = Field(default=[ExportFormat.all], description="Formats to export to")
    output_dir: Optional[str] = None


class GenerateFromPromptRequest(BaseModel):
    """Request model for natural language document generation."""
    prompt: str = Field(..., description="Natural language description of the document to generate")
    filename: Optional[str] = None
    theme: Optional[str] = "corporate"
    metadata: Optional[DocumentMetadata] = None


# ── Excel Models ──

class ExcelSheetData(BaseModel):
    """Data for a single Excel sheet."""
    name: str = Field(default="Sheet1", description="Sheet name (max 31 chars)")
    headers: list[str] = Field(default_factory=list, description="Column headers")
    rows: list[list[Any]] = Field(default_factory=list, description="Data rows")


class ExcelChartRequest(BaseModel):
    """Request to add a chart to an Excel sheet."""
    sheet_name: str = Field(..., description="Target sheet name")
    chart_type: ChartType = Field(default=ChartType.bar, description="Chart type")
    data_range: str = Field(..., description="Cell range for chart data (e.g., 'A1:B10')")
    title: Optional[str] = None
    position: Optional[str] = Field(default="D2", description="Cell position for chart")


class ExcelStyleRequest(BaseModel):
    """Request to apply styles to a cell range."""
    sheet_name: str = Field(..., description="Target sheet name")
    cell_range: str = Field(..., description="Cell range (e.g., 'A1:C10')")
    font_name: Optional[str] = None
    font_size: Optional[int] = None
    bold: Optional[bool] = None
    italic: Optional[bool] = None
    font_color: Optional[str] = None
    fill_color: Optional[str] = None
    border_color: Optional[str] = None
    border_style: Optional[str] = None
    alignment: Optional[str] = None
    number_format: Optional[str] = None
    text_wrap: Optional[bool] = None


class ExcelConditionalFormatRequest(BaseModel):
    """Request to add conditional formatting."""
    sheet_name: str = Field(..., description="Target sheet name")
    cell_range: str = Field(..., description="Cell range to apply formatting")
    rule_type: str = Field(..., description="Rule type: greater_than, less_than, between, equals, contains, data_bar, color_scale, icon_set")
    rule_value: Optional[Any] = None
    rule_value2: Optional[Any] = None
    format: Optional[dict] = None


class ExcelPivotRequest(BaseModel):
    """Request to create a pivot table."""
    sheet_name: str = Field(..., description="Source data sheet name")
    data_range: str = Field(..., description="Data range for pivot")
    row_fields: list[str] = Field(..., description="Fields for rows")
    value_fields: list[str] = Field(..., description="Fields for values")
    agg_function: str = Field(default="sum", description="Aggregation: sum, avg, count, min, max")
    output_sheet: Optional[str] = None


class ExcelDataValidationRequest(BaseModel):
    """Request to add data validation."""
    sheet_name: str = Field(..., description="Target sheet name")
    cell_range: str = Field(..., description="Cell range for validation")
    validation_type: str = Field(..., description="Type: list, whole, decimal, text_length")
    allowed_values: Optional[list[str]] = None
    operator: Optional[str] = None
    formula1: Optional[str] = None


class ExcelCreateRequest(BaseModel):
    """Request to create an Excel workbook."""
    filename: str = Field(..., description="Output filename")
    sheets: list[ExcelSheetData] = Field(..., description="Sheet data")
    theme: Optional[str] = "corporate"
    session_id: Optional[str] = None
    metadata: Optional[DocumentMetadata] = None
    locale: Optional[str] = "en_US"


# ── PowerPoint Models ──

class PPTXSlideRequest(BaseModel):
    """Request to add a slide."""
    layout: SlideLayout = Field(default=SlideLayout.title_and_content, description="Slide layout")
    title: Optional[str] = None
    content: Optional[str] = None
    bullets: Optional[list[str]] = None
    image_path: Optional[str] = None


class PPTXChartRequest(BaseModel):
    """Request to add a chart to a slide."""
    slide_index: int = Field(..., description="Slide index (0-based)")
    chart_type: ChartType = Field(default=ChartType.bar, description="Chart type")
    headers: list[str] = Field(..., description="Data headers")
    rows: list[list[Any]] = Field(..., description="Data rows")
    title: Optional[str] = None


class PPTXTableRequest(BaseModel):
    """Request to add a table to a slide."""
    slide_index: int = Field(..., description="Slide index (0-based)")
    headers: list[str] = Field(..., description="Table headers")
    rows: list[list[Any]] = Field(..., description="Table rows")


class PPTXTextBoxRequest(BaseModel):
    """Request to add a text box."""
    slide_index: int = Field(..., description="Slide index (0-based)")
    text: str = Field(..., description="Text content")
    left: float = Field(default=1.0, description="Left position in inches")
    top: float = Field(default=1.0, description="Top position in inches")
    width: float = Field(default=5.0, description="Width in inches")
    height: float = Field(default=1.0, description="Height in inches")
    font_size: Optional[int] = None
    bold: Optional[bool] = None
    color: Optional[str] = None


class PPTXImageRequest(BaseModel):
    """Request to add an image."""
    slide_index: int = Field(..., description="Slide index (0-based)")
    image_path: str = Field(..., description="Path to image file")
    left: float = Field(default=1.0, description="Left position in inches")
    top: float = Field(default=1.0, description="Top position in inches")
    width: Optional[float] = None
    height: Optional[float] = None


class PPTXCreateRequest(BaseModel):
    """Request to create a PowerPoint presentation."""
    filename: str = Field(..., description="Output filename")
    theme: Optional[str] = "corporate"
    title: Optional[str] = "Presentation"
    slides: list[PPTXSlideRequest] = Field(default_factory=list, description="Slides to add")
    slide_size: Optional[str] = "widescreen"
    session_id: Optional[str] = None
    metadata: Optional[DocumentMetadata] = None


# ── Word Models ──

class DOCXHeadingRequest(BaseModel):
    """Request to add a heading."""
    text: str = Field(..., description="Heading text")
    level: int = Field(default=1, ge=1, le=4, description="Heading level (1-4)")


class DOCXParagraphRequest(BaseModel):
    """Request to add a paragraph."""
    text: str = Field(..., description="Paragraph text")
    bold: Optional[bool] = None
    italic: Optional[bool] = None
    underline: Optional[bool] = None
    strikethrough: Optional[bool] = None
    font_name: Optional[str] = None
    font_size: Optional[int] = None
    color: Optional[str] = None
    alignment: Optional[str] = None
    space_before: Optional[int] = None
    space_after: Optional[int] = None


class DOCXTableRequest(BaseModel):
    """Request to add a table."""
    headers: list[str] = Field(..., description="Table headers")
    rows: list[list[Any]] = Field(..., description="Table rows")
    style: Optional[str] = "Table Grid"


class DOCXListRequest(BaseModel):
    """Request to add a list."""
    items: list[str] = Field(..., description="List items")
    ordered: bool = Field(default=False, description="True for numbered list, False for bulleted")


class DOCXImageRequest(BaseModel):
    """Request to add an image."""
    image_path: str = Field(..., description="Path to image file")
    caption: Optional[str] = None
    width: Optional[float] = None
    height: Optional[float] = None


class DOCXEquationRequest(BaseModel):
    """Request to add an equation."""
    equation: str = Field(..., description="Equation in LaTeX or OMML format")


class DOCXFooterHeaderRequest(BaseModel):
    """Request to add header/footer."""
    text: str = Field(..., description="Header/footer text")
    include_page_number: bool = Field(default=True, description="Include page numbers")
    include_date: bool = Field(default=False, description="Include date")


class DOCXCreateRequest(BaseModel):
    """Request to create a Word document."""
    filename: str = Field(..., description="Output filename")
    theme: Optional[str] = "corporate"
    title: Optional[str] = "Document"
    page_size: Optional[str] = "A4"
    orientation: Optional[str] = "portrait"
    session_id: Optional[str] = None
    metadata: Optional[DocumentMetadata] = None
    locale: Optional[str] = "en_US"