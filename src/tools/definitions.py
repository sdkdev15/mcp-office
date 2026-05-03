"""MCP Tool definitions for the office document generation server."""

from __future__ import annotations

from mcp.types import Tool


TOOLS = [
    # ── Excel ──
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
                                "items": {"type": "array", "items": {}, "description": "Each row is an array of cell values"},
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
    # NOTE: _from_prompt is listed FIRST — it should be the default choice.
    Tool(
        name="docx_from_prompt",
        description=(
            "PREFERRED tool for creating Word documents. Generate a complete, "
            "professionally formatted .docx file from a natural language prompt. "
            "Use this whenever the user asks you to create, write, draft, summarize, "
            "or generate any document — including resumes, reports, summaries, letters, "
            "essays, articles, meeting notes, proposals, etc. Simply pass the full "
            "content or instructions as the 'prompt' parameter. No need to structure "
            "paragraphs or tables yourself. This tool handles all formatting automatically."
        ),
        inputSchema={
            "type": "object",
            "properties": {
                "prompt": {
                    "type": "string",
                    "description": (
                        "The full content, text, or instructions to generate the document from. "
                        "This can be: the user's request verbatim, article text to summarize, "
                        "resume details, report content, or any natural language description. "
                        "Include all the information you want in the document."
                    ),
                },
                "filename": {
                    "type": "string",
                    "description": "Output filename (e.g., 'resume.docx', 'summary.docx'). Defaults to 'document.docx'.",
                },
                "theme": {"type": "string", "default": "corporate"},
                "session_id": {"type": "string"},
            },
            "required": ["prompt"],
        },
    ),
    Tool(
        name="docx_create",
        description=(
            "Create a Word document (.docx) from PRE-STRUCTURED data only. "
            "Requires explicit content_paragraphs array and/or tables array with "
            "headers/rows already prepared. Do NOT use this for natural language "
            "requests — use 'docx_from_prompt' instead, which is simpler "
            "and handles formatting automatically."
        ),
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
                                "items": {"type": "array", "items": {}, "description": "Each row is an array of cell values"},
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
    # NOTE: _from_prompt is listed FIRST — it should be the default choice.
    Tool(
        name="pptx_from_prompt",
        description=(
            "PREFERRED tool for creating PowerPoint presentations. Generate a "
            "complete, professionally formatted .pptx file from a natural language "
            "prompt. Use this whenever the user asks you to create, make, draft, or "
            "generate any presentation from a description, topic, source text, or "
            "instructions. Simply pass the full content as the 'prompt' parameter — "
            "no need to structure slides yourself. This tool handles all slide "
            "layout and formatting automatically."
        ),
        inputSchema={
            "type": "object",
            "properties": {
                "prompt": {
                    "type": "string",
                    "description": (
                        "The full content, text, or instructions to generate the presentation from. "
                        "This can be: the user's request verbatim, topic to present about, "
                        "article text to turn into slides, or any natural language description. "
                        "Include all the information you want in the presentation."
                    ),
                },
                "filename": {
                    "type": "string",
                    "description": "Output filename (e.g., 'quarterly_review.pptx'). Defaults to 'presentation.pptx'.",
                },
                "theme": {"type": "string", "default": "corporate"},
                "session_id": {"type": "string"},
            },
            "required": ["prompt"],
        },
    ),
    Tool(
        name="pptx_create",
        description=(
            "Create a PowerPoint presentation (.pptx) from PRE-STRUCTURED slide "
            "data only. Requires an explicit 'slides' array with slide objects "
            "containing titles, bullets, tables already prepared. Do NOT use this "
            "for natural language requests — use 'pptx_from_prompt' instead, "
            "which is simpler and handles slide layout automatically."
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
