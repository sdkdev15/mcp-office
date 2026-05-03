"""MCP Tool definitions for the office document generation server."""

from __future__ import annotations

from mcp.types import Tool


TOOLS = [
    Tool(
        name="excel_create",
        description="Create an Excel workbook (.xlsx) with multiple sheets, styling, and formatting.",
        inputSchema={
            "type": "object",
            "properties": {
                "filename": {"type": "string", "description": "Output filename (e.g., 'report.xlsx')"},
                "sheets": {
                    "type": "array",
                    "description": "List of sheet objects. Each sheet should have 'name' (string), 'headers' (array of strings), and 'rows' (array of arrays).",
                    "items": {
                        "type": "object",
                        "properties": {
                            "name": {"type": "string", "description": "Sheet name"},
                            "headers": {"type": "array", "items": {"type": "string"}, "description": "Column headers"},
                            "rows": {
                                "type": "array",
                                "items": {"type": "array", "description": "Each row is an array of cell values"},
                                "description": "Data rows"
                            },
                        },
                        "required": ["name", "headers", "rows"],
                    },
                },
                "theme": {"type": "string", "description": "Theme name (corporate, minimal, creative, academic, dark)", "default": "corporate"},
                "template_path": {"type": "string", "description": "Optional path to a .xlsx template file to use as base (preserves formatting, formulas, charts)"},
                "session_id": {"type": "string", "description": "Optional session ID"},
                "metadata": {"type": "object", "description": "Optional document metadata"},
                "locale": {"type": "string", "description": "Locale (en_US, id_ID)", "default": "en_US"},
            },
            "required": ["filename", "sheets"],
        },
    ),
    Tool(
        name="excel_export",
        description="Export data to multiple formats (xlsx, ods, csv).",
        inputSchema={
            "type": "object",
            "properties": {
                "sheets": {
                    "type": "array",
                    "description": "List of sheet objects with name, headers, and rows",
                    "items": {
                        "type": "object",
                        "properties": {
                            "name": {"type": "string", "description": "Sheet name"},
                            "headers": {"type": "array", "items": {"type": "string"}, "description": "Column headers"},
                            "rows": {
                                "type": "array",
                                "items": {"type": "array", "description": "Each row is an array of cell values"},
                                "description": "Data rows"
                            },
                        },
                        "required": ["name", "headers", "rows"],
                    },
                },
                "format": {"type": "string", "description": "xlsx, ods, csv, all", "default": "all"},
                "session_id": {"type": "string"},
            },
            "required": ["sheets"],
        },
    ),
    Tool(
        name="docx_create",
        description="Create a Word document (.docx) with page setup and theme styling.",
        inputSchema={
            "type": "object",
            "properties": {
                "filename": {"type": "string"},
                "title": {"type": "string", "default": "Document"},
                "theme": {"type": "string", "default": "corporate"},
                "template_path": {"type": "string", "description": "Optional path to a .docx template file to use as base (preserves formatting, styles, headers, footers)"},
                "page_size": {"type": "string", "default": "A4"},
                "orientation": {"type": "string", "default": "portrait"},
                "content_paragraphs": {"type": "array", "description": "Optional list of paragraph texts to append", "items": {"type": "string"}},
                "tables": {
                    "type": "array",
                    "description": "Optional list of table data with headers and rows",
                    "items": {
                        "type": "object",
                        "properties": {
                            "headers": {"type": "array", "items": {"type": "string"}, "description": "Table column headers"},
                            "rows": {
                                "type": "array",
                                "items": {"type": "array", "description": "Each row is an array of cell values"},
                                "description": "Table data rows"
                            },
                        },
                        "required": ["headers", "rows"],
                    },
                },
                "session_id": {"type": "string"},
                "metadata": {"type": "object"},
            },
            "required": ["filename"],
        },
    ),
    Tool(
        name="docx_generate_from_prompt",
        description="Generate a Word document from a natural language description.",
        inputSchema={
            "type": "object",
            "properties": {
                "prompt": {"type": "string"},
                "filename": {"type": "string"},
                "theme": {"type": "string", "default": "corporate"},
                "session_id": {"type": "string"},
            },
            "required": ["prompt"],
        },
    ),
    Tool(
        name="docx_export",
        description="Export a document in multiple formats (docx, odt).",
        inputSchema={
            "type": "object",
            "properties": {
                "title": {"type": "string"},
                "format": {"type": "string", "default": "all"},
                "theme": {"type": "string", "default": "corporate"},
                "session_id": {"type": "string"},
                "metadata": {"type": "object"},
            },
            "required": ["title"],
        },
    ),
    Tool(
        name="pptx_create",
        description="Create a PowerPoint presentation (.pptx) with slides and theme styling. Use this for structured slide data with titles, bullets, tables.",
        inputSchema={
            "type": "object",
            "additionalProperties": True,
            "properties": {
                "filename": {"type": "string", "description": "Output filename (e.g., 'presentation.pptx')"},
                "title": {"type": "string", "description": "Presentation title", "default": "Presentation"},
                "theme": {"type": "string", "description": "Theme name (corporate, minimal, creative, academic, dark)", "default": "corporate"},
                "template_path": {"type": "string", "description": "Optional path to a .pptx template file to use as base (preserves master slides, themes, layouts)"},
                "slides": {
                    "type": "array",
                    "description": "List of slide objects. Each slide can have: title (string), content (string), bullets (array of strings), table_headers (array of strings), table_rows (array of arrays), layout (string), image_path (string).",
                    "items": {
                        "type": "object",
                        "additionalProperties": True,
                        "properties": {
                            "title": {"type": "string", "description": "Slide title"},
                            "content": {"type": "string", "description": "Slide body content/description"},
                            "bullets": {"type": "array", "items": {"type": "string"}, "description": "Bullet points for the slide"},
                            "layout": {"type": "string", "description": "Slide layout (title_and_content, title, blank, two_content)", "default": "title_and_content"},
                            "image_path": {"type": "string", "description": "Optional path to an image"},
                            "table_headers": {"type": "array", "items": {"type": "string"}, "description": "Optional table column headers"},
                            "table_rows": {"type": "array", "items": {"type": "array", "items": {}}, "description": "Optional table data rows (array of arrays of cell values)"},
                        }
                    }
                },
                "slide_size": {"type": "string", "description": "Slide size (widescreen for 16:9, standard for 4:3)", "default": "widescreen"},
                "session_id": {"type": "string", "description": "Optional session ID"},
                "metadata": {"type": "object", "description": "Optional presentation metadata (author, company, subject, title, keywords, category, comments)"},
            },
            "required": ["filename"],
        },
    ),
    Tool(
        name="pptx_generate_from_prompt",
        description="Generate a PowerPoint presentation from a natural language description.",
        inputSchema={
            "type": "object",
            "properties": {
                "prompt": {"type": "string"},
                "filename": {"type": "string"},
                "theme": {"type": "string", "default": "corporate"},
                "session_id": {"type": "string"},
            },
            "required": ["prompt"],
        },
    ),
    Tool(
        name="pptx_export",
        description="Export a presentation in multiple formats (pptx, odp).",
        inputSchema={
            "type": "object",
            "properties": {
                "title": {"type": "string"},
                "format": {"type": "string", "default": "all"},
                "theme": {"type": "string", "default": "corporate"},
                "session_id": {"type": "string"},
                "metadata": {"type": "object"},
            },
            "required": ["title"],
        },
    ),
    Tool(
        name="list_themes_tool",
        description="List all available themes for document generation.",
        inputSchema={"type": "object", "properties": {}},
    ),
    Tool(
        name="list_files",
        description="List all generated files for a session.",
        inputSchema={
            "type": "object",
            "properties": {"session_id": {"type": "string"}},
            "required": ["session_id"],
        },
    ),
    Tool(
        name="get_storage_stats",
        description="Get current storage usage statistics.",
        inputSchema={"type": "object", "properties": {}},
    ),
]
