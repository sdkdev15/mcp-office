"""Unit tests for MCP Office Server."""

import pytest
import io
import json
from pathlib import Path
from datetime import datetime, timedelta

from src.utils.validators import (
    validate_string,
    validate_filename,
    validate_sheet_data,
    validate_chart_type,
    validate_slide_layout,
    validate_heading_level,
    validate_export_format,
    validate_table_data,
    ValidationError,
)
from src.utils.data_transformer import DataTransformer
from src.utils.rate_limiter import RateLimiter
from src.utils.file_handler import FileHandler
from src.utils.security import PIIRedactor, InputSanitizer, AuditTrail
from src.styles.themes import get_theme, list_themes, THEMES


# ── Validator Tests ──

class TestValidators:
    def test_validate_string_required(self):
        with pytest.raises(ValidationError):
            validate_string(None, "name", required=True)

    def test_validate_string_min_length(self):
        with pytest.raises(ValidationError):
            validate_string("ab", "name", min_length=3, required=True)

    def test_validate_string_max_length(self):
        with pytest.raises(ValidationError):
            validate_string("a" * 1001, "name", max_length=1000)

    def test_validate_string_success(self):
        assert validate_string("hello", "name") == "hello"

    def test_validate_filename_dangerous_chars(self):
        result = validate_filename("test/file.xlsx")
        assert "/" not in result

    def test_validate_filename_empty(self):
        with pytest.raises(ValidationError):
            validate_filename("")

    def test_validate_sheet_data_empty(self):
        with pytest.raises(ValidationError):
            validate_sheet_data([])

    def test_validate_sheet_data_duplicate_names(self):
        with pytest.raises(ValidationError):
            validate_sheet_data([
                {"name": "Sheet1", "headers": [], "rows": []},
                {"name": "Sheet1", "headers": [], "rows": []},
            ])

    def test_validate_chart_type_invalid(self):
        with pytest.raises(ValidationError):
            validate_chart_type("invalid_type")

    def test_validate_chart_type_success(self):
        assert validate_chart_type("bar") == "bar"

    def test_validate_slide_layout_success(self):
        assert validate_slide_layout("title_and_content") == "title_and_content"

    def test_validate_heading_level_invalid(self):
        with pytest.raises(ValidationError):
            validate_heading_level(5)

    def test_validate_heading_level_success(self):
        assert validate_heading_level(1) == 1

    def test_validate_export_format_success(self):
        assert validate_export_format("xlsx") == "xlsx"


# ── Data Transformer Tests ──

class TestDataTransformer:
    def setup_method(self):
        self.transformer = DataTransformer()

    def test_json_to_table_dict_list(self):
        data = [
            {"name": "Alice", "age": 30},
            {"name": "Bob", "age": 25},
        ]
        result = self.transformer.json_to_table_data(data)
        assert result["headers"] == ["name", "age"]
        assert len(result["rows"]) == 2

    def test_json_to_table_list_list(self):
        data = [
            ["Name", "Age"],
            ["Alice", 30],
            ["Bob", 25],
        ]
        result = self.transformer.json_to_table_data(data)
        assert result["headers"] == ["Name", "Age"]
        assert len(result["rows"]) == 2

    def test_csv_to_table_data(self):
        csv_str = "Name,Age\nAlice,30\nBob,25"
        result = self.transformer.csv_to_table_data(csv_str)
        assert result["headers"] == ["Name", "Age"]
        assert len(result["rows"]) == 2

    def test_detect_chart_type_time_series(self):
        headers = ["Month", "Sales"]
        rows = [["Jan", 100], ["Feb", 200], ["Mar", 150]]
        assert self.transformer.detect_chart_type(headers, rows) == "line"

    def test_detect_chart_type_bar(self):
        headers = ["Category", "Value"]
        rows = [["A", 10], ["B", 20], ["C", 30]]
        assert self.transformer.detect_chart_type(headers, rows) == "bar"

    def test_format_number_indonesian(self):
        result = self.transformer.format_number(1234.5, "id_ID")
        assert "," in result  # Decimal comma in Indonesian

    def test_format_currency_idr(self):
        result = self.transformer.format_currency(1000000, "IDR", "id_ID")
        assert "Rp" in result

    def test_flatten_data(self):
        data = {"a": {"b": 1}, "c": [2, 3]}
        result = self.transformer.flatten_data(data)
        assert "a.b" in result

    def test_merge_tables(self):
        tables = [
            {"headers": ["A", "B"], "rows": [[1, 2]]},
            {"headers": ["A", "B"], "rows": [[3, 4]]},
        ]
        result = self.transformer.merge_tables(tables)
        assert len(result["rows"]) == 2


# ── Rate Limiter Tests ──

class TestRateLimiter:
    def test_rate_limiter_allows_requests(self):
        limiter = RateLimiter(max_requests=5, window_seconds=60)
        allowed, info = limiter.is_allowed("user1")
        assert allowed is True
        assert info["remaining"] == 4

    def test_rate_limiter_blocks_when_exceeded(self):
        limiter = RateLimiter(max_requests=2, window_seconds=60)
        limiter.is_allowed("user1")
        limiter.is_allowed("user1")
        allowed, info = limiter.is_allowed("user1")
        assert allowed is False
        assert "retry_after" in info

    def test_rate_limiter_different_users(self):
        limiter = RateLimiter(max_requests=1, window_seconds=60)
        allowed1, _ = limiter.is_allowed("user1")
        allowed2, _ = limiter.is_allowed("user2")
        assert allowed1 is True
        assert allowed2 is True

    def test_rate_limiter_reset(self):
        limiter = RateLimiter(max_requests=1, window_seconds=60)
        limiter.is_allowed("user1")
        limiter.reset_user("user1")
        allowed, _ = limiter.is_allowed("user1")
        assert allowed is True


# ── PII Redactor Tests ──

class TestPIIRedactor:
    def setup_method(self):
        self.redactor = PIIRedactor()

    def test_redact_email(self):
        result = self.redactor.redact("Contact: john@example.com")
        assert "[EMAIL]" in result
        assert "john@example.com" not in result

    def test_redact_phone(self):
        result = self.redactor.redact("Phone: +6281234567890")
        assert "[PHONE]" in result

    def test_redact_indonesian_phone(self):
        result = self.redactor.redact("HP: 081234567890")
        assert "[PHONE]" in result

    def test_redact_data_structure(self):
        data = {
            "name": "John",
            "email": "john@example.com",
            "details": ["info@test.com", "safe text"],
        }
        result = self.redactor.redact_data(data)
        assert "[EMAIL]" in result["email"]
        assert "[EMAIL]" in result["details"][0]
        assert "safe text" in result["details"][1]


# ── Input Sanitizer Tests ──

class TestInputSanitizer:
    def setup_method(self):
        self.sanitizer = InputSanitizer()

    def test_sanitize_script_tags(self):
        result = self.sanitizer.sanitize("<script>alert('xss')</script>Hello")
        assert "<script>" not in result

    def test_sanitize_path_traversal(self):
        result = self.sanitizer.sanitize("../../../etc/passwd")
        assert "../" not in result

    def test_is_safe(self):
        assert self.sanitizer.is_safe("Normal text here") is True
        assert self.sanitizer.is_safe("<script>alert(1)</script>") is False


# ── Theme Tests ──

class TestThemes:
    def test_get_theme_exists(self):
        theme = get_theme("corporate")
        assert theme.name == "corporate"
        assert theme.colors.primary == "#1E40AF"

    def test_get_theme_fallback(self):
        theme = get_theme("nonexistent")
        assert theme.name == "corporate"  # Falls back to corporate

    def test_list_themes(self):
        themes = list_themes()
        assert len(themes) == 5
        assert all("name" in t for t in themes)
        assert all("description" in t for t in themes)

    def test_theme_colors(self):
        for name in THEMES:
            theme = get_theme(name)
            assert hasattr(theme, "colors")
            assert hasattr(theme, "fonts")
            assert theme.colors.primary.startswith("#")


# ── File Handler Tests ──

class TestFileHandler:
    def test_generate_filename(self):
        handler = FileHandler()
        filename = handler.generate_filename("test file", ".xlsx")
        assert filename.endswith(".xlsx")
        assert " " not in filename

    def test_generate_filename_unique(self):
        handler = FileHandler()
        f1 = handler.generate_filename("test", ".xlsx")
        f2 = handler.generate_filename("test", ".xlsx")
        assert f1 != f2  # Unique IDs ensure different filenames


# ── Shared Utility Tests ──

class TestColors:
    def test_hex_to_rgbcolor_tuple_with_hash(self):
        from src.utils.colors import hex_to_rgbcolor_tuple
        assert hex_to_rgbcolor_tuple("#1E40AF") == (30, 64, 175)

    def test_hex_to_rgbcolor_tuple_without_hash(self):
        from src.utils.colors import hex_to_rgbcolor_tuple
        assert hex_to_rgbcolor_tuple("FF0000") == (255, 0, 0)

    def test_ensure_argb_hex_6char(self):
        from src.utils.colors import ensure_argb_hex
        assert ensure_argb_hex("#1E40AF") == "FF1E40AF"

    def test_ensure_argb_hex_8char(self):
        from src.utils.colors import ensure_argb_hex
        assert ensure_argb_hex("FF1E40AF") == "FF1E40AF"

    def test_ensure_argb_hex_empty(self):
        from src.utils.colors import ensure_argb_hex
        assert ensure_argb_hex("") == "FF000000"


class TestFormatting:
    def test_human_readable_bytes(self):
        from src.utils.formatting import human_readable_size
        assert "B" in human_readable_size(500)

    def test_human_readable_kb(self):
        from src.utils.formatting import human_readable_size
        assert "KB" in human_readable_size(2048)

    def test_human_readable_mb(self):
        from src.utils.formatting import human_readable_size
        assert "MB" in human_readable_size(5 * 1024 * 1024)


class TestMetadata:
    def test_apply_metadata_author(self):
        from src.utils.metadata import apply_metadata

        class FakeProps:
            author = None
            title = None
        props = FakeProps()
        apply_metadata(props, {"author": "Test Author", "title": "Test Title"})
        assert props.author == "Test Author"
        assert props.title == "Test Title"

    def test_apply_metadata_creator_fallback(self):
        from src.utils.metadata import apply_metadata

        class FakeProps:
            creator = None  # openpyxl uses 'creator' not 'author'
        props = FakeProps()
        apply_metadata(props, {"author": "Creator Test"})
        assert props.creator == "Creator Test"

    def test_apply_metadata_skips_empty(self):
        from src.utils.metadata import apply_metadata

        class FakeProps:
            author = "Original"
        props = FakeProps()
        apply_metadata(props, {"author": "", "unknown_field": "ignored"})
        assert props.author == "Original"


# ── Generator Output Tests ──

class TestExcelGenerator:
    def test_create_workbook_returns_bytes(self):
        from src.generators.excel_generator import ExcelGenerator
        gen = ExcelGenerator("corporate")
        sheets = [{"name": "Test", "headers": ["A", "B"], "rows": [[1, 2]]}]
        data = gen.create_workbook(sheets)
        assert isinstance(data, bytes)
        assert len(data) > 0
        # XLSX magic bytes (PK zip header)
        assert data[:2] == b"PK"

    def test_create_workbook_with_metadata(self):
        from src.generators.excel_generator import ExcelGenerator
        gen = ExcelGenerator("minimal")
        sheets = [{"name": "Sheet1", "headers": ["X"], "rows": [["val"]]}]
        data = gen.create_workbook(sheets, metadata={"author": "Test"})
        assert isinstance(data, bytes)
        assert len(data) > 0


class TestDOCXGenerator:
    def test_create_document_returns_bytes(self):
        from src.generators.docx_generator import DOCXGenerator
        gen = DOCXGenerator("corporate")
        data = gen.create_document("Test Doc")
        assert isinstance(data, bytes)
        assert len(data) > 0
        assert data[:2] == b"PK"

    def test_create_document_with_content(self):
        from src.generators.docx_generator import DOCXGenerator
        gen = DOCXGenerator("corporate")
        data = gen.create_document_with_content(
            title="Test",
            content_paragraphs=["Hello world", "Second paragraph"],
            tables=[{"headers": ["A", "B"], "rows": [["1", "2"]]}],
        )
        assert isinstance(data, bytes)
        assert len(data) > 0

    def test_create_from_prompt(self):
        from src.generators.docx_generator import DOCXGenerator
        gen = DOCXGenerator("corporate")
        data = gen.create_from_prompt("Write a short report about testing")
        assert isinstance(data, bytes)
        assert len(data) > 0


class TestPPTXGenerator:
    def test_create_presentation_returns_bytes(self):
        from src.generators.pptx_generator import PPTXGenerator
        gen = PPTXGenerator("corporate")
        data = gen.create_presentation("Test Pres")
        assert isinstance(data, bytes)
        assert len(data) > 0
        assert data[:2] == b"PK"

    def test_add_slide_with_content(self):
        """Regression test: placeholder_type AttributeError fix."""
        from src.generators.pptx_generator import PPTXGenerator
        gen = PPTXGenerator("corporate")
        gen.create_presentation("Test")
        # This used to fail with 'SlidePlaceholder has no attribute placeholder_type'
        idx = gen.add_slide(
            title="Content Slide",
            content="Body text here",
            layout="title_and_content",
        )
        assert idx >= 0

    def test_add_slide_with_bullets(self):
        """Regression test: bullet formatting via placeholder."""
        from src.generators.pptx_generator import PPTXGenerator
        gen = PPTXGenerator("corporate")
        gen.create_presentation("Test")
        idx = gen.add_slide(
            title="Bullet Slide",
            bullets=["Point 1", "Point 2", "Point 3"],
            layout="title_and_content",
        )
        assert idx >= 0

    def test_create_from_prompt(self):
        from src.generators.pptx_generator import PPTXGenerator
        gen = PPTXGenerator("corporate")
        data = gen.create_from_prompt("Create a 3 slide presentation about Python")
        assert isinstance(data, bytes)
        assert len(data) > 0


if __name__ == "__main__":
    pytest.main([__file__, "-v"])