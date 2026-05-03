"""Pydantic models for tool input/output schemas."""

from src.models.schemas import (
    ExcelCreateRequest,
    ExcelSheetData,
    ExcelChartRequest,
    ExcelStyleRequest,
    ExcelConditionalFormatRequest,
    ExcelPivotRequest,
    ExcelDataValidationRequest,
    PPTXCreateRequest,
    PPTXSlideRequest,
    PPTXChartRequest,
    PPTXTableRequest,
    PPTXTextBoxRequest,
    PPTXImageRequest,
    DOCXCreateRequest,
    DOCXHeadingRequest,
    DOCXParagraphRequest,
    DOCXTableRequest,
    DOCXListRequest,
    DOCXImageRequest,
    DOCXEquationRequest,
    DOCXFooterHeaderRequest,
    ExportRequest,
    GenerateFromPromptRequest,
    DocumentMetadata,
)

__all__ = [
    # Excel
    "ExcelCreateRequest",
    "ExcelSheetData",
    "ExcelChartRequest",
    "ExcelStyleRequest",
    "ExcelConditionalFormatRequest",
    "ExcelPivotRequest",
    "ExcelDataValidationRequest",
    # PowerPoint
    "PPTXCreateRequest",
    "PPTXSlideRequest",
    "PPTXChartRequest",
    "PPTXTableRequest",
    "PPTXTextBoxRequest",
    "PPTXImageRequest",
    # Word
    "DOCXCreateRequest",
    "DOCXHeadingRequest",
    "DOCXParagraphRequest",
    "DOCXTableRequest",
    "DOCXListRequest",
    "DOCXImageRequest",
    "DOCXEquationRequest",
    "DOCXFooterHeaderRequest",
    # Common
    "ExportRequest",
    "GenerateFromPromptRequest",
    "DocumentMetadata",
]