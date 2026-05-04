"""MCP Tool definitions for the office document generation server."""

from __future__ import annotations

from mcp.types import Tool


TOOLS = [
    # ── Excel ──
    Tool(
        name="excel_create",
        description=(
            "Create an Excel workbook (.xlsx) with multiple sheets, styling, charts, and formulas.\n\n"
            "SHEETS JSON SAMPLE:\n"
            "[\n"
            "  {\n"
            "    \"name\": \"Revenue\",\n"
            "    \"headers\": [\"Month\", \"Sales\", \"Expenses\", \"Profit\"],\n"
            "    \"rows\": [\n"
            "      [\"Jan\", 50000, 30000, \"=B2-C2\"],\n"
            "      [\"Feb\", 62000, 35000, \"=B3-C3\"],\n"
            "      [\"Total\", \"=SUM(B2:B3)\", \"=SUM(C2:C3)\", \"=SUM(D2:D3)\"]\n"
            "    ],\n"
            "    \"charts\": [{\"chart_type\": \"bar\", \"data_range\": \"A1:B3\", \"title\": \"Monthly Sales\", \"position\": \"F2\"}]\n"
            "  }\n"
            "]\n\n"
            "TIPS: Cell values starting with '=' are treated as Excel formulas. "
            "Supported chart types: bar, line, pie, column, area."
        ),
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
                                "items": {"type": "array", "items": {}, "description": "Each row is an array of cell values"},
                                "description": "Data rows"
                            },
                            "charts": {
                                "type": "array",
                                "description": "Optional list of charts to generate",
                                "items": {
                                    "type": "object",
                                    "properties": {
                                        "chart_type": {"type": "string", "description": "bar, line, pie, column, area"},
                                        "data_range": {"type": "string", "description": "Cell range for data (e.g., 'A1:B10')"},
                                        "title": {"type": "string", "description": "Chart title", "default": "Chart"},
                                        "position": {"type": "string", "description": "Placement (e.g., 'D2')", "default": "D2"}
                                    },
                                    "required": ["chart_type", "data_range"]
                                }
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
                                "items": {"type": "array", "items": {}, "description": "Each row is an array of cell values"},
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

    # ── Word Documents (DOCX) ──
    Tool(
        name="docx_create",
        description=(
            "PREFERRED tool for creating Word documents (.docx). Use a structured 'sections' array to build "
            "rich documents with titles, subtitles, table of contents, headings (H1-H3), paragraphs, bullet lists, "
            "numbered lists, and tables — all interleaved in any order.\n\n"
            "SECTIONS JSON SAMPLE:\n"
            "[\n"
            "  {\"type\": \"title\", \"text\": \"Quarterly Report\"},\n"
            "  {\"type\": \"subtitle\", \"text\": \"Q3 2026 Financial Summary\"},\n"
            "  {\"type\": \"toc\"},\n"
            "  {\"type\": \"heading_1\", \"text\": \"1. Revenue Analysis\"},\n"
            "  {\"type\": \"paragraph\", \"text\": \"Revenue grew by 12% year-over-year.\"},\n"
            "  {\"type\": \"list_bullet\", \"items\": [\"SaaS: +15%\", \"Services: +8%\"]},\n"
            "  {\"type\": \"heading_2\", \"text\": \"1.1 Regional Breakdown\"},\n"
            "  {\"type\": \"list_number\", \"items\": [\"APAC grew fastest\", \"EMEA stable\"]},\n"
            "  {\"type\": \"table\", \"headers\": [\"Region\", \"Revenue\"], \"rows\": [[\"APAC\", \"$2.1M\"]]}\n"
            "]"
        ),
        inputSchema={
            "type": "object",
            "properties": {
                "filename": {"type": "string", "description": "Output filename (e.g., 'report.docx')"},
                "title": {"type": "string", "description": "Document title (used as fallback if no title section is provided)", "default": "Document"},
                "theme": {"type": "string", "description": "Theme name (corporate, minimal, creative, academic, dark)", "default": "corporate"},
                "template_path": {"type": "string", "description": "Optional path to a .docx template file to use as base (preserves formatting, styles, headers, footers)"},
                "page_size": {"type": "string", "default": "A4"},
                "orientation": {"type": "string", "default": "portrait"},
                "sections": {
                    "type": "array",
                    "description": (
                        "Ordered array of document sections. Each section has a 'type' and type-specific fields. "
                        "Supported types: title, subtitle, toc, heading_1, heading_2, heading_3, paragraph, list_bullet, list_number, table."
                    ),
                    "items": {
                        "type": "object",
                        "additionalProperties": True,
                        "properties": {
                            "type": {
                                "type": "string",
                                "enum": ["title", "subtitle", "toc", "heading_1", "heading_2", "heading_3", "paragraph", "list_bullet", "list_number", "table"],
                                "description": "The section element type",
                            },
                            "text": {"type": "string", "description": "Text content (for title, subtitle, heading_*, paragraph)"},
                            "items": {"type": "array", "items": {"type": "string"}, "description": "List items (for list_bullet, list_number)"},
                            "headers": {"type": "array", "items": {"type": "string"}, "description": "Table column headers (for table)"},
                            "rows": {"type": "array", "items": {"type": "array", "items": {}}, "description": "Table data rows (for table)"},
                        },
                        "required": ["type"],
                    },
                },
                "content_paragraphs": {"type": "array", "description": "DEPRECATED: Use 'sections' instead. Legacy list of paragraph texts.", "items": {"type": "string"}},
                "tables": {
                    "type": "array",
                    "description": "DEPRECATED: Use 'sections' with type 'table' instead. Legacy table data.",
                    "items": {
                        "type": "object",
                        "properties": {
                            "headers": {"type": "array", "items": {"type": "string"}},
                            "rows": {"type": "array", "items": {"type": "array", "items": {}}},
                        },
                        "required": ["headers", "rows"],
                    },
                },
                "session_id": {"type": "string"},
                "metadata": {"type": "object", "description": "Optional document metadata (author, company, subject, title, keywords, category, comments)"},
            },
            "required": ["filename"],
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

    # ── PowerPoint Presentations (PPTX) ──
    Tool(
        name="pptx_create",
        description=(
            "PREFERRED tool for creating PowerPoint presentations (.pptx). "
            "Pass an explicit 'slides' JSON array with slide objects. Each slide supports: "
            "title, content, bullets, tables, and images.\n\n"
            "SLIDES JSON SAMPLE:\n"
            "[\n"
            "  {\"title\": \"Welcome\", \"layout\": \"title\"},\n"
            "  {\"title\": \"Overview\", \"content\": \"Key highlights of Q3\", \"bullets\": [\"Revenue up 12%\", \"New markets opened\"], \"layout\": \"title_and_content\"},\n"
            "  {\"title\": \"Financials\", \"layout\": \"title_and_content\", \"table_headers\": [\"Metric\", \"Value\"], \"table_rows\": [[\"Revenue\", \"$2.1M\"], [\"Profit\", \"$400K\"]]},\n"
            "  {\"title\": \"Thank You\", \"content\": \"Questions?\", \"layout\": \"title_only\"}\n"
            "]\n\n"
            "AVAILABLE LAYOUTS: title, title_and_content, title_only, two_content, blank, section_header"
        ),
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

    # ── Utility Tools ──
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
