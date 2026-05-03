"""Data transformation utilities for converting JSON/CSV to document data structures."""

from __future__ import annotations

import csv
import io
import json
from typing import Any, Optional

from src.utils.logger import get_logger

log = get_logger("data_transformer")


class DataTransformer:
    """Transforms various data formats into document-ready structures."""

    @staticmethod
    def json_to_table_data(data: Any) -> dict[str, Any]:
        """Convert JSON data to table format (headers + rows).

        Supports:
        - List of dicts: keys become headers
        - List of lists: first row becomes headers
        - Single dict: converted to key-value table

        Args:
            data: JSON data to convert.

        Returns:
            Dictionary with 'headers' and 'rows' keys.
        """
        if isinstance(data, dict):
            return {
                "headers": ["Key", "Value"],
                "rows": [[str(k), str(v)] for k, v in data.items()],
            }

        if isinstance(data, list) and len(data) == 0:
            return {"headers": [], "rows": []}

        if isinstance(data, list) and isinstance(data[0], dict):
            headers = list(data[0].keys())
            rows = [[row.get(h, "") for h in headers] for row in data]
            return {"headers": headers, "rows": rows}

        if isinstance(data, list) and isinstance(data[0], list):
            return {
                "headers": [str(h) for h in data[0]],
                "rows": [[str(cell) for cell in row] for row in data[1:]],
            }

        return {"headers": ["Value"], "rows": [[str(data)]]}

    @staticmethod
    def csv_to_table_data(csv_string: str) -> dict[str, Any]:
        """Convert CSV string to table format.

        Args:
            csv_string: CSV formatted string.

        Returns:
            Dictionary with 'headers' and 'rows' keys.
        """
        reader = csv.reader(io.StringIO(csv_string))
        rows = list(reader)

        if not rows:
            return {"headers": [], "rows": []}

        return {
            "headers": [str(h) for h in rows[0]],
            "rows": [[str(cell) for cell in row] for row in rows[1:]],
        }

    @staticmethod
    def detect_chart_type(headers: list[str], rows: list[list]) -> str:
        """Auto-detect the best chart type based on data patterns.

        Heuristics:
        - Time series data (date-like first column) -> line
        - Two columns (category + value) -> bar
        - Percentage/sum-to-100 data -> pie
        - Multiple value columns -> column
        - Three numeric columns -> scatter

        Args:
            headers: Column headers.
            rows: Data rows.

        Returns:
            Recommended chart type string.
        """
        if not rows or not headers:
            return "bar"

        num_columns = len(headers)
        first_col = [str(row[0]).lower() for row in rows[:10]]

        # Detect time series
        time_indicators = ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec",
                           "mon", "tue", "wed", "thu", "fri", "sat", "sun",
                           "q1", "q2", "q3", "q4", "year", "month", "week", "day"]
        is_time_series = any(
            any(ind in cell for ind in time_indicators)
            for cell in first_col
        )

        # Count numeric columns
        numeric_cols = 0
        for col_idx in range(num_columns):
            is_numeric = False
            for row in rows[:5]:
                if col_idx < len(row):
                    try:
                        float(str(row[col_idx]))
                        is_numeric = True
                    except (ValueError, TypeError):
                        break
            if is_numeric:
                numeric_cols += 1

        # Detection logic
        if is_time_series:
            return "line"

        if numeric_cols == 1 and num_columns == 2:
            return "bar"

        # Check if data sums to ~100 (percentage data -> pie)
        if numeric_cols == 1:
            values = []
            for row in rows:
                for cell in row:
                    try:
                        values.append(float(cell))
                    except (ValueError, TypeError):
                        pass
            if values and 95 <= sum(values) <= 105:
                return "pie"

        if numeric_cols >= 3:
            return "scatter"

        if numeric_cols >= 2:
            return "column"

        return "bar"

    @staticmethod
    def format_number(value: Any, locale: str = "en_US") -> str:
        """Format a number according to locale settings.

        Args:
            value: Number to format.
            locale: Locale string (e.g., 'id_ID', 'en_US').

        Returns:
            Formatted number string.
        """
        try:
            num = float(value)
        except (ValueError, TypeError):
            return str(value)

        if locale == "id_ID":
            # Indonesian: comma as decimal, dot as thousands
            return f"{num:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        elif locale == "en_US":
            return f"{num:,.2f}"
        else:
            return f"{num:,.2f}"

    @staticmethod
    def format_currency(value: Any, currency: str = "USD", locale: str = "en_US") -> str:
        """Format a number as currency.

        Args:
            value: Number to format.
            currency: Currency code.
            locale: Locale string.

        Returns:
            Formatted currency string.
        """
        try:
            num = float(value)
        except (ValueError, TypeError):
            return str(value)

        symbols = {
            "USD": "$",
            "EUR": "€",
            "GBP": "£",
            "IDR": "Rp",
            "JPY": "¥",
            "CNY": "¥",
        }

        symbol = symbols.get(currency.upper(), currency)
        formatted = DataTransformer.format_number(num, locale)

        if currency.upper() == "IDR":
            return f"{symbol} {formatted.replace('.00', '')}"
        return f"{symbol}{formatted}"

    @staticmethod
    def flatten_data(data: Any, prefix: str = "") -> dict:
        """Flatten nested data structures.

        Args:
            data: Data to flatten.
            prefix: Key prefix for nested keys.

        Returns:
            Flattened dictionary.
        """
        result = {}

        if isinstance(data, dict):
            for key, value in data.items():
                new_key = f"{prefix}.{key}" if prefix else key
                if isinstance(value, (dict, list)):
                    result.update(DataTransformer.flatten_data(value, new_key))
                else:
                    result[new_key] = value
        elif isinstance(data, list):
            for i, item in enumerate(data):
                new_key = f"{prefix}.{i}" if prefix else str(i)
                if isinstance(item, (dict, list)):
                    result.update(DataTransformer.flatten_data(item, new_key))
                else:
                    result[new_key] = item
        else:
            result[prefix] = data

        return result

    @staticmethod
    def merge_tables(tables: list[dict]) -> dict[str, Any]:
        """Merge multiple table data structures.

        Args:
            tables: List of table dictionaries with 'headers' and 'rows'.

        Returns:
            Merged table dictionary.
        """
        if not tables:
            return {"headers": [], "rows": []}

        if len(tables) == 1:
            return tables[0]

        # Use first table's headers, union all headers
        all_headers = list(tables[0]["headers"])
        for table in tables[1:]:
            for header in table["headers"]:
                if header not in all_headers:
                    all_headers.append(header)

        # Merge rows
        all_rows = []
        for table in tables:
            for row in table.get("rows", []):
                # Pad row to match all_headers length
                merged_row = list(row)
                while len(merged_row) < len(all_headers):
                    merged_row.append("")
                all_rows.append(merged_row[:len(all_headers)])

        return {"headers": all_headers, "rows": all_rows}


# Global transformer instance
data_transformer = DataTransformer()