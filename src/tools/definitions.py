"""MCP Tool definitions for the office document generation server."""

from __future__ import annotations

from mcp.types import Tool


TOOLS = [
    # ── Excel ──
    Tool(
        name="excel_create",
        description=(
            "Create an Excel workbook (.xlsx) with multiple sheets, styling, charts, and formulas.\n\n"
            "Use EITHER 'sheets' (new) OR 'template_path' (from template), not both.\n\n"
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
            "Supported chart types: bar, line, pie, column, area, doughnut, radar, scatter, bubble."
        ),
        inputSchema={
            "type": "object",
            "properties": {
                "filename": {
                    "type": "string",
                    "pattern": r"^[\w\-]+\.xlsx?$",
                    "description": "Output filename (e.g., 'report.xlsx')"
                },
                "sheets": {
                    "type": "array",
                    "description": "List of sheet objects. Required if template_path not provided.",
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
                                        "chart_type": {
                                            "type": "string",
                                            "enum": ["bar", "line", "pie", "column", "area", "doughnut", "radar", "scatter", "bubble"],
                                            "description": "Chart type"
                                        },
                                        "data_range": {
                                            "type": "string",
                                            "pattern": r"^[A-Z]+\d+:[A-Z]+\d+$",
                                            "description": "Cell range for data (e.g., 'A1:B10')"
                                        },
                                        "title": {"type": "string", "description": "Chart title", "default": "Chart"},
                                        "position": {
                                            "type": "string",
                                            "pattern": r"^[A-Z]+\d+$",
                                            "description": "Placement (e.g., 'D2')",
                                            "default": "D2"
                                        }
                                    },
                                    "required": ["chart_type", "data_range"]
                                }
                            },
                        },
                        "required": ["name", "headers", "rows"],
                    },
                },
                "theme": {
                    "type": "string",
                    "enum": ["corporate", "minimal", "creative", "academic", "dark"],
                    "description": "Theme name for styling",
                    "default": "corporate"
                },
                "template_path": {
                    "type": "string",
                    "pattern": r"^[\w\-./]+\.xlsx$",
                    "description": "Optional path to a .xlsx template file to use as base (preserves formatting, formulas, charts). Cannot be used with 'sheets'."
                },
                "session_id": {"type": "string", "description": "Optional session ID"},
                "metadata": {"type": "object", "description": "Optional document metadata"},
                "locale": {
                    "type": "string",
                    "enum": ["en_US", "id_ID"],
                    "description": "Locale (en_US, id_ID)",
                    "default": "en_US"
                },
            },
            "required": ["filename"],
            "examples": [
                {
                    "filename": "sales_report.xlsx",
                    "theme": "corporate",
                    "sheets": [
                        {
                            "name": "Q1",
                            "headers": ["Jan", "Feb", "Mar"],
                            "rows": [[100, 200, 150]]
                        }
                    ]
                }
            ]
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
                "format": {
                    "type": "string",
                    "enum": ["xlsx", "ods", "csv", "all"],
                    "description": "Output format",
                    "default": "all"
                },
                "session_id": {"type": "string"},
            },
            "required": ["sheets"],
            "examples": [
                {
                    "format": "csv",
                    "sheets": [
                        {
                            "name": "Export Data",
                            "headers": ["ID", "Value"],
                            "rows": [[1, 100], [2, 200]]
                        }
                    ]
                }
            ]
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
                "filename": {
                    "type": "string",
                    "pattern": r"^[\w\-]+\.docx?$",
                    "description": "Output filename (e.g., 'report.docx')"
                },
                "title": {"type": "string", "description": "Document title (used as fallback if no title section is provided)", "default": "Document"},
                "theme": {
                    "type": "string",
                    "enum": ["corporate", "minimal", "creative", "academic", "dark"],
                    "description": "Theme name",
                    "default": "corporate"
                },
                "template_path": {
                    "type": "string",
                    "pattern": r"^[\w\-./]+\.docx$",
                    "description": "Optional path to a .docx template file to use as base (preserves formatting, styles, headers, footers). Cannot be used with 'sections'."
                },
                "page_size": {
                    "type": "string",
                    "enum": ["A4", "Letter", "Legal"],
                    "default": "A4"
                },
                "orientation": {
                    "type": "string",
                    "enum": ["portrait", "landscape"],
                    "default": "portrait"
                },
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
            "examples": [
                {
                    "filename": "summary.docx",
                    "title": "Meeting Summary",
                    "sections": [
                        {"type": "heading_1", "text": "Overview"},
                        {"type": "paragraph", "text": "This is a summary of the meeting."}
                    ]
                }
            ]
        },
    ),
    Tool(
        name="docx_export",
        description="Export a document in multiple formats (docx, odt).",
        inputSchema={
            "type": "object",
            "properties": {
                "title": {"type": "string"},
                "format": {
                    "type": "string",
                    "enum": ["docx", "odt", "all"],
                    "default": "all"
                },
                "theme": {
                    "type": "string",
                    "enum": ["corporate", "minimal", "creative", "academic", "dark"],
                    "default": "corporate"
                },
                "session_id": {"type": "string"},
                "metadata": {"type": "object"},
            },
            "required": ["title"],
            "examples": [
                {
                    "title": "My Export Document",
                    "format": "odt"
                }
            ]
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
            "AVAILABLE LAYOUTS: title, title_and_content, title_only, two_content, blank, section_header, comparison"
        ),
        inputSchema={
            "type": "object",
            "additionalProperties": True,
            "properties": {
                "filename": {
                    "type": "string",
                    "pattern": r"^[\w\-]+\.pptx?$",
                    "description": "Output filename (e.g., 'presentation.pptx')"
                },
                "title": {"type": "string", "description": "Presentation title", "default": "Presentation"},
                "theme": {
                    "type": "string",
                    "enum": ["corporate", "minimal", "creative", "academic", "dark"],
                    "description": "Theme name",
                    "default": "corporate"
                },
                "template_path": {
                    "type": "string",
                    "pattern": r"^[\w\-./]+\.pptx$",
                    "description": "Optional path to a .pptx template file to use as base (preserves master slides, themes, layouts)"
                },
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
                            "layout": {
                                "type": "string",
                                "enum": ["title", "title_and_content", "title_only", "two_content", "blank", "section_header", "comparison"],
                                "description": "Slide layout",
                                "default": "title_and_content"
                            },
                            "image_path": {"type": "string", "description": "Optional path to an image"},
                            "table_headers": {"type": "array", "items": {"type": "string"}, "description": "Optional table column headers"},
                            "table_rows": {"type": "array", "items": {"type": "array", "items": {}}, "description": "Optional table data rows (array of arrays of cell values)"},
                        }
                    }
                },
                "slide_size": {
                    "type": "string",
                    "enum": ["widescreen", "standard"],
                    "description": "Slide size (widescreen for 16:9, standard for 4:3)",
                    "default": "widescreen"
                },
                "session_id": {"type": "string", "description": "Optional session ID"},
                "metadata": {"type": "object", "description": "Optional presentation metadata (author, company, subject, title, keywords, category, comments)"},
            },
            "required": ["filename"],
            "examples": [
                {
                    "filename": "deck.pptx",
                    "title": "Quarterly Update",
                    "slides": [
                        {"title": "Q1 Performance", "layout": "title"}
                    ]
                }
            ]
        },
    ),
    Tool(
        name="pptx_export",
        description="Export a presentation in multiple formats (pptx, odp).",
        inputSchema={
            "type": "object",
            "properties": {
                "title": {"type": "string"},
                "format": {
                    "type": "string",
                    "enum": ["pptx", "odp", "all"],
                    "default": "all"
                },
                "theme": {
                    "type": "string",
                    "enum": ["corporate", "minimal", "creative", "academic", "dark"],
                    "default": "corporate"
                },
                "session_id": {"type": "string"},
                "metadata": {"type": "object"},
            },
            "required": ["title"],
            "examples": [
                {
                    "title": "My Presentation",
                    "format": "odp"
                }
            ]
        },
    ),

    # ── Analysis & Generation ──
    Tool(
        name="analyze_data",
        description="Analyze a dataset and return statistics, trends, and correlations.",
        inputSchema={
            "type": "object",
            "properties": {
                "data": {
                    "type": "array",
                    "description": "Array of JSON objects representing the dataset rows",
                    "items": {"type": "object"}
                },
                "target_columns": {
                    "type": "array",
                    "items": {"type": "string"},
                    "description": "Optional list of numeric columns to analyze"
                },
                "breakdown_by": {
                    "type": "string",
                    "description": "Optional column name to group data by (e.g., 'month', 'category')"
                }
            },
            "required": ["data"]
        }
    ),
    Tool(
        name="generate_summary",
        description="Generate high-level text summaries, key metrics, and insights based on raw data.",
        inputSchema={
            "type": "object",
            "properties": {
                "data": {
                    "type": "array",
                    "description": "Array of JSON objects representing the dataset rows",
                    "items": {"type": "object"}
                },
                "style": {
                    "type": "string",
                    "enum": ["professional", "casual", "technical"],
                    "default": "professional"
                },
                "include_metrics": {"type": "boolean", "default": True},
                "max_insights": {"type": "integer", "default": 5}
            },
            "required": ["data"]
        }
    ),
    Tool(
        name="generate_faq",
        description="Formulate Q&A pairs based on statistical analysis of the dataset.",
        inputSchema={
            "type": "object",
            "properties": {
                "data": {
                    "type": "array",
                    "description": "Array of JSON objects representing the dataset rows",
                    "items": {"type": "object"}
                },
                "num_questions": {"type": "integer", "default": 5},
                "question_style": {
                    "type": "string",
                    "enum": ["practical", "formal"],
                    "default": "practical"
                }
            },
            "required": ["data"]
        }
    ),
    Tool(
        name="recommend_charts",
        description="Evaluate dataset dimensions and recommend optimal chart types.",
        inputSchema={
            "type": "object",
            "properties": {
                "data": {
                    "type": "array",
                    "description": "Array of JSON objects representing the dataset rows",
                    "items": {"type": "object"}
                },
                "data_types": {
                    "type": "object",
                    "description": "Optional mapping of column name to type ('numeric', 'categorical', 'time')"
                },
                "num_recommendations": {"type": "integer", "default": 3}
            },
            "required": ["data"]
        }
    ),

    # ── Advanced Visualization & Formatting (Track C) ──
    Tool(
        name="excel_advanced_formatting",
        description="Create an Excel file with conditional formatting (data bars, color scales, cell rules).",
        inputSchema={
            "type": "object",
            "properties": {
                "filename": {"type": "string", "pattern": r"^[\w\-. ]+\.xlsx$"},
                "theme": {"type": "string", "enum": ["corporate", "minimal", "creative", "academic", "dark"]},
                "sheets": {
                    "type": "array",
                    "items": {
                        "type": "object",
                        "properties": {
                            "name": {"type": "string"},
                            "headers": {"type": "array", "items": {"type": "string"}},
                            "rows": {"type": "array", "items": {"type": ["string", "number", "boolean"]}},
                            "conditional_formatting": {
                                "type": "array",
                                "items": {
                                    "type": "object",
                                    "properties": {
                                        "data_range": {"type": "string", "description": "e.g., 'B2:B10'"},
                                        "type": {"type": "string", "enum": ["data_bar", "color_scale", "cell_is"]},
                                        "color": {"type": "string", "description": "Hex color for data_bar"},
                                        "start_color": {"type": "string"},
                                        "mid_color": {"type": "string"},
                                        "end_color": {"type": "string"},
                                        "operator": {"type": "string", "enum": ["greaterThan", "lessThan", "equal", "between"]},
                                        "formula": {"type": "array", "items": {"type": "string"}},
                                        "fill_color": {"type": "string"},
                                        "font_color": {"type": "string"}
                                    }
                                }
                            }
                        }
                    }
                }
            },
            "required": ["filename", "sheets"]
        }
    ),
    Tool(
        name="excel_with_images",
        description="Create an Excel file containing embedded images (from URLs or Base64).",
        inputSchema={
            "type": "object",
            "properties": {
                "filename": {"type": "string", "pattern": r"^[\w\-. ]+\.xlsx$"},
                "sheets": {
                    "type": "array",
                    "items": {
                        "type": "object",
                        "properties": {
                            "name": {"type": "string"},
                            "images": {
                                "type": "array",
                                "items": {
                                    "type": "object",
                                    "properties": {
                                        "source": {"type": "string", "description": "URL or Base64 data:image/png;base64,... string"},
                                        "position": {"type": "string", "description": "Cell reference like 'E5'"},
                                        "width": {"type": "integer"},
                                        "height": {"type": "integer"}
                                    },
                                    "required": ["source"]
                                }
                            }
                        }
                    }
                }
            },
            "required": ["filename", "sheets"]
        }
    ),
    Tool(
        name="docx_with_images",
        description="Create a Word document containing inline or floating images.",
        inputSchema={
            "type": "object",
            "properties": {
                "filename": {"type": "string", "pattern": r"^[\w\-. ]+\.docx$"},
                "sections": {
                    "type": "array",
                    "items": {
                        "type": "object",
                        "properties": {
                            "type": {"type": "string", "enum": ["image", "paragraph", "heading_1", "heading_2"]},
                            "source": {"type": "string", "description": "For image type: URL or Base64"},
                            "width": {"type": "number", "description": "Width in inches"},
                            "caption": {"type": "string"},
                            "text": {"type": "string", "description": "For text types"}
                        },
                        "required": ["type"]
                    }
                }
            },
            "required": ["filename", "sections"]
        }
    ),

    # ── Batch & Templating (Track D) ──
    Tool(
        name="batch_create_documents",
        description="Generate multiple documents from a single template and a list of datasets (mail merge). Returns a ZIP file.",
        inputSchema={
            "type": "object",
            "properties": {
                "format": {"type": "string", "enum": ["excel", "docx", "pptx"]},
                "template": {
                    "type": "string",
                    "description": "JSON string representing the document payload. Use {{var}} for variables, {{#each array}}...{{/each}} for loops, {{#if cond}}...{{/if}} for conditionals."
                },
                "datasets": {
                    "type": "array",
                    "description": "Array of data objects to inject into the template. Each object produces one document.",
                    "items": {"type": "object"}
                },
                "theme": {"type": "string", "enum": ["corporate", "minimal", "creative", "academic", "dark"], "default": "corporate"}
            },
            "required": ["format", "template", "datasets"]
        }
    ),
    Tool(
        name="merge_documents",
        description="Merge multiple Word (.docx) or PDF (.pdf) documents into a single file.",
        inputSchema={
            "type": "object",
            "properties": {
                "file_paths": {
                    "type": "array",
                    "description": "Array of absolute file paths to merge.",
                    "items": {"type": "string"}
                },
                "output_filename": {
                    "type": "string",
                    "description": "Filename for the merged output file (must end in .docx or .pdf)"
                }
            },
            "required": ["file_paths", "output_filename"]
        }
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
