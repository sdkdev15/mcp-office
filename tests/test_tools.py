import pytest
from src.tools.definitions import TOOLS

def test_tool_schemas():
    for tool in TOOLS:
        assert tool.name
        assert tool.description
        assert "inputSchema" in tool.model_dump()
        schema = tool.inputSchema
        
        # Tools with examples should have 'examples' list
        if "examples" in schema:
            assert isinstance(schema["examples"], list)
            assert len(schema["examples"]) > 0
            
        # Ensure 'theme' has an enum if it exists
        if "theme" in schema["properties"]:
            assert "enum" in schema["properties"]["theme"]
            assert "corporate" in schema["properties"]["theme"]["enum"]

def test_excel_create_schema():
    excel_tool = next(t for t in TOOLS if t.name == "excel_create")
    schema = excel_tool.inputSchema
    assert "filename" in schema["properties"]
    assert "pattern" in schema["properties"]["filename"]
    
    assert "examples" in schema
    assert schema["examples"][0]["filename"] == "sales_report.xlsx"

def test_mutually_exclusive_docs():
    excel_tool = next(t for t in TOOLS if t.name == "excel_create")
    desc = excel_tool.description
    assert "mutually exclusive" in desc.lower() or "not both" in desc.lower()
