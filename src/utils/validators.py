"""Input validation utilities for document generation tools."""

from __future__ import annotations

import re
import difflib
from typing import Any

from src.utils.logger import get_logger

log = get_logger("validators")

# Maximum limits for validation
MAX_SHEET_NAME_LENGTH = 31
MAX_CELL_VALUE_LENGTH = 32767
MAX_ROWS_PER_SHEET = 1_048_576
MAX_COLUMNS_PER_SHEET = 16_384
MAX_SLIDES = 500
MAX_PAGES = 500
MAX_TABLE_ROWS = 10_000
MAX_TABLE_COLUMNS = 100
MAX_FILE_SIZE_MB = 50
ALLOWED_CHART_TYPES = [
    "bar", "column", "line", "pie", "doughnut", "area", "radar", "scatter", "bubble"
]
ALLOWED_LAYOUTS = [
    "title", "title_and_content", "blank", "two_content", "comparison",
    "section_header", "title_only"
]
ALLOWED_HEADING_LEVELS = [1, 2, 3, 4]
ALLOWED_EXPORT_FORMATS = ["xlsx", "ods", "csv", "pptx", "odp", "docx", "odt", "pdf", "all"]
ALLOWED_THEMES = ["corporate", "minimal", "creative", "academic", "dark"]
ALLOWED_LOCALES = ["en_US", "id_ID"]


class ValidationError(Exception):
    """Custom validation error."""

    def __init__(self, field: str, message: str):
        self.field = field
        self.message = message
        super().__init__(f"Validation error in '{field}': {message}")


def suggest_closest(word: str, possibilities: list[str]) -> str:
    """Return a suggestion string if a close match is found."""
    matches = difflib.get_close_matches(word, possibilities, n=1, cutoff=0.6)
    return f" Did you mean '{matches[0]}'?" if matches else ""


def validate_mutually_exclusive(data: dict, fields: list[str]):
    """Ensure that only one of the mutually exclusive fields is provided."""
    provided = [f for f in fields if data.get(f) is not None]
    if len(provided) > 1:
        raise ValidationError(
            ", ".join(provided),
            f"These parameters are mutually exclusive. Provide only one of: {', '.join(fields)}"
        )


def validate_string(value: Any, field: str, min_length: int = 0, max_length: int = 1000, required: bool = False) -> str:
    """Validate and sanitize a string value."""
    if value is None:
        if required:
            raise ValidationError(field, "This field is required")
        return ""

    if not isinstance(value, str):
        value = str(value)

    if len(value) < min_length:
        raise ValidationError(field, f"Minimum length is {min_length}, got {len(value)}")

    if len(value) > max_length:
        raise ValidationError(field, f"Maximum length is {max_length}, got {len(value)}")

    return value.strip()


def validate_filename(filename: str) -> str:
    """Validate and sanitize a filename."""
    if not filename:
        raise ValidationError("filename", "Filename is required")

    # Check for invalid characters using regex
    if re.search(r'[<>:"/\\|?*]', filename):
        raise ValidationError("filename", "Filename contains invalid characters (<>:/\\|?*)")

    import os.path
    allowed_extensions = {".xlsx", ".ods", ".csv", ".pptx", ".odp", ".docx", ".odt", ".pdf"}
    _, ext = os.path.splitext(filename.lower())
    if ext and ext not in allowed_extensions:
        raise ValidationError("filename", f"Invalid extension '{ext}'. Allowed: {', '.join(allowed_extensions)}")

    if len(filename) > 200:
        raise ValidationError("filename", "Filename too long (max 200 characters)")

    return filename


def validate_cell_range(cell_range: str, field: str) -> str:
    """Validate Excel cell range like 'A1:B10'."""
    cell_range = str(cell_range).upper().strip()
    if not re.match(r"^[A-Z]+\d+:[A-Z]+\d+$", cell_range):
        raise ValidationError(field, f"Invalid cell range format '{cell_range}'. Expected format like 'A1:B10'")
    return cell_range


def validate_formula(formula: str, field: str) -> str:
    """Validate Excel formula syntax (starts with '=')."""
    formula = str(formula).strip()
    if not formula.startswith("="):
        raise ValidationError(field, f"Formulas must start with '='. Did you mean '={formula}'?")
    return formula


def validate_sheet_data(sheets: list[dict]) -> list[dict]:
    """Validate Excel sheet data structure."""
    if not sheets:
        raise ValidationError("sheets", "At least one sheet is required")

    if len(sheets) > 50:
        raise ValidationError("sheets", "Maximum 50 sheets allowed")

    sheet_names = set()
    for i, sheet in enumerate(sheets):
        name = sheet.get("name", f"Sheet{i + 1}")

        if len(name) > MAX_SHEET_NAME_LENGTH:
            raise ValidationError(f"sheets[{i}].name", f"Sheet name max length is {MAX_SHEET_NAME_LENGTH}")

        if name in sheet_names:
            raise ValidationError(f"sheets[{i}].name", f"Duplicate sheet name: {name}")
        sheet_names.add(name)

        headers = sheet.get("headers", [])
        if headers and not isinstance(headers, list):
            raise ValidationError(f"sheets[{i}].headers", "Headers must be a list")

        rows = sheet.get("rows", [])
        if not isinstance(rows, list):
            raise ValidationError(f"sheets[{i}].rows", "Rows must be a list")

        if len(rows) > MAX_ROWS_PER_SHEET:
            raise ValidationError(f"sheets[{i}].rows", f"Maximum {MAX_ROWS_PER_SHEET} rows allowed")

        if headers:
            for j, row in enumerate(rows):
                if not isinstance(row, list):
                    raise ValidationError(f"sheets[{i}].rows[{j}]", "Row must be a list")
                if len(row) != len(headers):
                    raise ValidationError(
                        f"sheets[{i}].rows[{j}]",
                        f"Row has {len(row)} columns, expected {len(headers)} columns based on headers. All rows must have the same length as headers."
                    )
                
                # Formula validation
                for k, cell in enumerate(row):
                    if isinstance(cell, str) and len(cell) > 1 and cell.startswith("SUM(") and not cell.startswith("="):
                         raise ValidationError(f"sheets[{i}].rows[{j}][{k}]", f"Formula missing '=' prefix: '{cell}'. Use '={cell}'.")
                    if isinstance(cell, str) and cell.startswith("="):
                         if len(cell) == 1:
                             raise ValidationError(f"sheets[{i}].rows[{j}][{k}]", "Empty formula.")

        # Chart range validation
        charts = sheet.get("charts", [])
        for k, chart in enumerate(charts):
            dr = chart.get("data_range")
            if dr:
                validate_cell_range(dr, f"sheets[{i}].charts[{k}].data_range")
            ctype = chart.get("chart_type")
            if ctype:
                validate_chart_type(ctype)

    return sheets


def validate_chart_type(chart_type: str) -> str:
    chart_type = chart_type.lower().strip()
    if chart_type not in ALLOWED_CHART_TYPES:
        suggestion = suggest_closest(chart_type, ALLOWED_CHART_TYPES)
        raise ValidationError(
            "chart_type",
            f"Invalid chart type '{chart_type}'. Allowed: {', '.join(ALLOWED_CHART_TYPES)}.{suggestion}"
        )
    return chart_type


def validate_slide_layout(layout: str) -> str:
    layout = layout.lower().strip().replace(" ", "_")
    if layout not in ALLOWED_LAYOUTS:
        suggestion = suggest_closest(layout, ALLOWED_LAYOUTS)
        raise ValidationError(
            "layout",
            f"Invalid layout '{layout}'. Allowed: {', '.join(ALLOWED_LAYOUTS)}.{suggestion}"
        )
    return layout


def validate_theme(theme: str) -> str:
    theme = theme.lower().strip()
    if theme not in ALLOWED_THEMES:
        suggestion = suggest_closest(theme, ALLOWED_THEMES)
        raise ValidationError(
            "theme",
            f"Invalid theme '{theme}'. Allowed: {', '.join(ALLOWED_THEMES)}.{suggestion}"
        )
    return theme


def validate_locale(locale: str) -> str:
    locale = locale.strip()
    if locale not in ALLOWED_LOCALES:
        suggestion = suggest_closest(locale, ALLOWED_LOCALES)
        raise ValidationError(
            "locale",
            f"Invalid locale '{locale}'. Allowed: {', '.join(ALLOWED_LOCALES)}.{suggestion}"
        )
    return locale


def validate_heading_level(level: int) -> int:
    if level not in ALLOWED_HEADING_LEVELS:
        raise ValidationError(
            "level",
            f"Invalid heading level {level}. Allowed: {ALLOWED_HEADING_LEVELS}"
        )
    return level


def validate_export_format(fmt: str) -> str:
    fmt = fmt.lower().strip()
    if fmt not in ALLOWED_EXPORT_FORMATS:
        suggestion = suggest_closest(fmt, ALLOWED_EXPORT_FORMATS)
        raise ValidationError(
            "format",
            f"Invalid export format '{fmt}'. Allowed: {', '.join(ALLOWED_EXPORT_FORMATS)}.{suggestion}"
        )
    return fmt


def validate_table_data(headers: list, rows: list) -> tuple[list, list]:
    if not headers:
        raise ValidationError("headers", "Table must have headers")

    if len(headers) > MAX_TABLE_COLUMNS:
        raise ValidationError("headers", f"Maximum {MAX_TABLE_COLUMNS} columns allowed")

    if len(rows) > MAX_TABLE_ROWS:
        raise ValidationError("rows", f"Maximum {MAX_TABLE_ROWS} rows allowed")

    for i, row in enumerate(rows):
        if not isinstance(row, list):
             raise ValidationError(f"rows[{i}]", "Row must be a list")
        if len(row) != len(headers):
             raise ValidationError(f"rows[{i}]", f"Row length ({len(row)}) does not match headers length ({len(headers)}).")

    return headers, rows


def validate_inputs(data: dict, schema: dict) -> dict:
    validated = {}
    for field, rules in schema.items():
        value = data.get(field)
        required = rules.get("required", False)
        field_type = rules.get("type", str)

        if value is None:
            if required:
                raise ValidationError(field, "This field is required")
            if "default" in rules:
                validated[field] = rules["default"]
            continue

        if field_type == str and not isinstance(value, str):
            validated[field] = str(value)
        elif field_type == int and not isinstance(value, int):
            try:
                validated[field] = int(value)
            except (ValueError, TypeError):
                raise ValidationError(field, f"Value must be an integer: {value}")
        elif field_type == list and not isinstance(value, list):
            raise ValidationError(field, f"Value must be a list: {value}")
        elif field_type == dict and not isinstance(value, dict):
            raise ValidationError(field, f"Value must be an object: {value}")
        else:
            validated[field] = value

    return validated