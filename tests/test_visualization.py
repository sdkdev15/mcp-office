import pytest
import os
from src.utils.image_handler import ImageHandler
from src.utils.conditional_formatting import ConditionalFormatter
from openpyxl import Workbook
import base64

def test_image_handler_base64(tmp_path):
    handler = ImageHandler(cache_dir=str(tmp_path))
    
    # Tiny 1x1 transparent PNG in base64
    b64_png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYAAAAAYAAjCB0C8AAAAASUVORK5CYII="
    data_uri = f"data:image/png;base64,{b64_png}"
    
    result_path = handler.process_image(data_uri)
    assert os.path.exists(result_path)
    assert result_path.endswith(".png")

def test_image_handler_invalid():
    handler = ImageHandler()
    with pytest.raises(ValueError):
        handler.process_image("invalid_source")

def test_conditional_formatter():
    wb = Workbook()
    ws = wb.active
    
    rules = [
        {"type": "data_bar", "data_range": "A1:A10", "color": "FF0000"},
        {"type": "color_scale", "data_range": "B1:B10"},
        {"type": "cell_is", "data_range": "C1:C10", "operator": "greaterThan", "formula": 50, "fill_color": "00FF00"}
    ]
    
    formatter = ConditionalFormatter()
    formatter.apply_rules(ws, rules)
    
    assert len(ws.conditional_formatting._cf_rules) == 3
    # Check that rules were added
    ranges = [str(r) for r in ws.conditional_formatting._cf_rules.keys()]
    assert any("A1:A10" in r for r in ranges)
    assert any("B1:B10" in r for r in ranges)
    assert any("C1:C10" in r for r in ranges)
