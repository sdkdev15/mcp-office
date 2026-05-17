import pytest
from src.utils.validators import (
    validate_string, validate_filename, validate_sheet_data, validate_chart_type,
    validate_slide_layout, validate_theme, validate_locale, validate_heading_level,
    validate_export_format, validate_table_data, validate_cell_range, validate_formula,
    validate_mutually_exclusive, ValidationError
)

def test_validate_string():
    assert validate_string(" test ", "field") == "test"
    with pytest.raises(ValidationError, match="This field is required"):
        validate_string(None, "field", required=True)
    with pytest.raises(ValidationError, match="Minimum length is 3"):
        validate_string("ab", "field", min_length=3)
    with pytest.raises(ValidationError, match="Maximum length is 5"):
        validate_string("abcdef", "field", max_length=5)

def test_validate_filename():
    assert validate_filename("report.xlsx") == "report.xlsx"
    with pytest.raises(ValidationError, match="Filename is required"):
        validate_filename("")
    with pytest.raises(ValidationError, match="invalid characters"):
        validate_filename("bad/name.xlsx")
    with pytest.raises(ValidationError, match="Invalid extension"):
        validate_filename("report.txt")

def test_validate_cell_range():
    assert validate_cell_range("A1:B10", "range") == "A1:B10"
    assert validate_cell_range(" aa1:Bb10 ", "range") == "AA1:BB10"
    with pytest.raises(ValidationError, match="Invalid cell range format"):
        validate_cell_range("A1-B10", "range")
    with pytest.raises(ValidationError, match="Invalid cell range format"):
        validate_cell_range("1A:10B", "range")

def test_validate_formula():
    assert validate_formula("=SUM(A1)", "form") == "=SUM(A1)"
    with pytest.raises(ValidationError, match="must start with '='"):
        validate_formula("SUM(A1)", "form")

def test_validate_sheet_data():
    sheets = [{"name": "S1", "headers": ["A", "B"], "rows": [[1, 2], [3, 4]]}]
    assert validate_sheet_data(sheets) == sheets

    with pytest.raises(ValidationError, match="At least one sheet is required"):
        validate_sheet_data([])

    with pytest.raises(ValidationError, match="Sheet name max length"):
        validate_sheet_data([{"name": "a" * 32, "headers": [], "rows": []}])

    with pytest.raises(ValidationError, match="Duplicate sheet name"):
        validate_sheet_data([{"name": "S1", "headers": [], "rows": []}, {"name": "S1", "headers": [], "rows": []}])

    with pytest.raises(ValidationError, match="expected 2 columns"):
        validate_sheet_data([{"name": "S1", "headers": ["A", "B"], "rows": [[1, 2], [3]]}])

    with pytest.raises(ValidationError, match="missing '=' prefix"):
        validate_sheet_data([{"name": "S1", "headers": ["A"], "rows": [["SUM(A1:B1)"]]}])

    with pytest.raises(ValidationError, match="Empty formula"):
        validate_sheet_data([{"name": "S1", "headers": ["A"], "rows": [["="]]}])

def test_validate_enums():
    # chart type
    assert validate_chart_type("bar") == "bar"
    with pytest.raises(ValidationError, match="Did you mean 'bar'?"):
        validate_chart_type("barr")

    # slide layout
    assert validate_slide_layout("title_and_content") == "title_and_content"
    with pytest.raises(ValidationError, match="Invalid layout"):
        validate_slide_layout("invalid")

    # theme
    assert validate_theme("corporate") == "corporate"
    with pytest.raises(ValidationError, match="Invalid theme"):
        validate_theme("blue")

    # locale
    assert validate_locale("en_US") == "en_US"
    with pytest.raises(ValidationError, match="Invalid locale"):
        validate_locale("en_UK")

def test_validate_heading_level():
    assert validate_heading_level(1) == 1
    with pytest.raises(ValidationError):
        validate_heading_level(5)

def test_validate_export_format():
    assert validate_export_format("xlsx") == "xlsx"
    with pytest.raises(ValidationError):
        validate_export_format("txt")

def test_validate_table_data():
    h, r = validate_table_data(["A", "B"], [[1, 2]])
    assert len(h) == 2
    assert len(r) == 1

    with pytest.raises(ValidationError, match="Table must have headers"):
        validate_table_data([], [[1, 2]])

    with pytest.raises(ValidationError, match="Row length"):
        validate_table_data(["A", "B"], [[1]])

def test_validate_mutually_exclusive():
    validate_mutually_exclusive({"a": 1}, ["a", "b"])
    with pytest.raises(ValidationError, match="mutually exclusive"):
        validate_mutually_exclusive({"a": 1, "b": 2}, ["a", "b"])
