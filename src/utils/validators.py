"""Input validation utilities for document generation tools."""

from __future__ import annotations

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


class ValidationError(Exception):
    """Custom validation error."""

    def __init__(self, field: str, message: str):
        self.field = field
        self.message = message
        super().__init__(f"Validation error in '{field}': {message}")


def validate_string(value: Any, field: str, min_length: int = 0, max_length: int = 1000, required: bool = False) -> str:
    """Validate and sanitize a string value.

    Args:
        value: Value to validate.
        field: Field name for error messages.
        min_length: Minimum string length.
        max_length: Maximum string length.
        required: Whether the field is required.

    Returns:
        Validated string.

    Raises:
        ValidationError: If validation fails.
    """
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
    """Validate and sanitize a filename.

    Args:
        filename: Filename to validate.

    Returns:
        Sanitized filename.

    Raises:
        ValidationError: If validation fails.
    """
    if not filename:
        raise ValidationError("filename", "Filename is required")

    # Remove dangerous characters
    dangerous_chars = ['/', '\\', ':', '*', '?', '"', '<', '>', '|']
    safe_name = filename
    for char in dangerous_chars:
        safe_name = safe_name.replace(char, "_")

    # Ensure it has a valid extension
    allowed_extensions = {".xlsx", ".ods", ".csv", ".pptx", ".odp", ".docx", ".odt", ".pdf"}
    _, ext = __import__("os.path").path.splitext(safe_name.lower())
    if ext and ext not in allowed_extensions:
        raise ValidationError("filename", f"Invalid extension '{ext}'. Allowed: {', '.join(allowed_extensions)}")

    if len(safe_name) > 200:
        raise ValidationError("filename", "Filename too long (max 200 characters)")

    return safe_name


def validate_sheet_data(sheets: list[dict]) -> list[dict]:
    """Validate Excel sheet data structure.

    Args:
        sheets: List of sheet data dictionaries.

    Returns:
        Validated sheet data.

    Raises:
        ValidationError: If validation fails.
    """
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

        # Validate headers
        headers = sheet.get("headers", [])
        if headers and not isinstance(headers, list):
            raise ValidationError(f"sheets[{i}].headers", "Headers must be a list")

        # Validate rows
        rows = sheet.get("rows", [])
        if not isinstance(rows, list):
            raise ValidationError(f"sheets[{i}].rows", "Rows must be a list")

        if len(rows) > MAX_ROWS_PER_SHEET:
            raise ValidationError(f"sheets[{i}].rows", f"Maximum {MAX_ROWS_PER_SHEET} rows allowed")

        if headers:
            for j, row in enumerate(rows):
                if isinstance(row, list) and len(row) != len(headers):
                    log.warning(f"Row {j} has {len(row)} columns, expected {len(headers)}")

    return sheets


def validate_chart_type(chart_type: str) -> str:
    """Validate chart type.

    Args:
        chart_type: Chart type string.

    Returns:
        Validated chart type.

    Raises:
        ValidationError: If chart type is invalid.
    """
    chart_type = chart_type.lower().strip()
    if chart_type not in ALLOWED_CHART_TYPES:
        raise ValidationError(
            "chart_type",
            f"Invalid chart type '{chart_type}'. Allowed: {', '.join(ALLOWED_CHART_TYPES)}"
        )
    return chart_type


def validate_slide_layout(layout: str) -> str:
    """Validate slide layout.

    Args:
        layout: Layout string.

    Returns:
        Validated layout.

    Raises:
        ValidationError: If layout is invalid.
    """
    layout = layout.lower().strip().replace(" ", "_")
    if layout not in ALLOWED_LAYOUTS:
        raise ValidationError(
            "layout",
            f"Invalid layout '{layout}'. Allowed: {', '.join(ALLOWED_LAYOUTS)}"
        )
    return layout


def validate_heading_level(level: int) -> int:
    """Validate heading level.

    Args:
        level: Heading level (1-4).

    Returns:
        Validated heading level.

    Raises:
        ValidationError: If level is invalid.
    """
    if level not in ALLOWED_HEADING_LEVELS:
        raise ValidationError(
            "level",
            f"Invalid heading level {level}. Allowed: {ALLOWED_HEADING_LEVELS}"
        )
    return level


def validate_export_format(fmt: str) -> str:
    """Validate export format.

    Args:
        fmt: Format string.

    Returns:
        Validated format.

    Raises:
        ValidationError: If format is invalid.
    """
    fmt = fmt.lower().strip()
    if fmt not in ALLOWED_EXPORT_FORMATS:
        raise ValidationError(
            "format",
            f"Invalid export format '{fmt}'. Allowed: {', '.join(ALLOWED_EXPORT_FORMATS)}"
        )
    return fmt


def validate_table_data(headers: list, rows: list) -> tuple[list, list]:
    """Validate table data structure.

    Args:
        headers: Table headers.
        rows: Table rows.

    Returns:
        Tuple of (headers, rows).

    Raises:
        ValidationError: If validation fails.
    """
    if not headers:
        raise ValidationError("headers", "Table must have headers")

    if len(headers) > MAX_TABLE_COLUMNS:
        raise ValidationError("headers", f"Maximum {MAX_TABLE_COLUMNS} columns allowed")

    if len(rows) > MAX_TABLE_ROWS:
        raise ValidationError("rows", f"Maximum {MAX_TABLE_ROWS} rows allowed")

    return headers, rows


def validate_inputs(data: dict, schema: dict) -> dict:
    """Validate input data against a schema definition.

    Args:
        data: Input data dictionary.
        schema: Schema definition with field requirements.

    Returns:
        Validated and sanitized data.

    Raises:
        ValidationError: If validation fails.
    """
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