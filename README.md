# MCP Office Server

MCP (Model Context Protocol) server for generating Office documents — **Excel (.xlsx)**, **Word (.docx)**, and **PowerPoint (.pptx)** — with ODF format support (.ods, .odt, .odp), 5 built-in themes, and Docker deployment.

## Features

- **Excel** — Multi-sheet workbooks with styling, charts, formulas, and conditional formatting
- **Word** — Structured documents with title, subtitle, TOC, headings (H1–H3), paragraphs, bullet/numbered lists, tables, and images
- **PowerPoint** — Presentations with corporate branding, slide layouts, tables, charts, and text boxes
- **ODF Support** — LibreOffice formats (.ods, .odt, .odp)
- **5 Built-in Themes** — corporate, minimal, creative, academic, dark
- **Custom Templates** — Use your own .xlsx/.docx/.pptx files as base templates
- **Multi-format Export** — Generate OOXML + ODF in a single call
- **Locale Support** — English (en_US) and Indonesian (id_ID)
- **Rate Limiting** — Sliding window per user
- **Auto Cleanup** — Automatic deletion of old files
- **Session Isolation** — Per-session output directories
- **Security** — PII redaction, input sanitization, audit trail
- **Docker Ready** — Build and run via Docker (stdio/SSE modes)

---

## Getting Started

### Desktop Mode (Python via stdio)

**1. Install dependencies:**

```bash
pip install openpyxl python-docx python-pptx odfpy mcp loguru pydantic pydantic-settings aiofiles aiohttp uvicorn starlette
```

**2. Configure MCP Client:**

```json
{
  "mcpServers": {
    "mcp-office": {
      "type": "stdio",
      "command": "python",
      "args": ["run_server.py"],
      "env": {
        "OUTPUT_DIR": "/path/to/mcp-office/outputs",
        "MCP_TRANSPORT": "stdio"
      }
    }
  }
}
```

### Docker Mode (SSE)

```bash
docker compose build
docker compose up -d
```

```json
{
  "mcpServers": {
    "mcp-office": {
      "type": "sse",
      "url": "http://YOUR_SERVER_IP:8765/sse"
    }
  }
}
```

---

## Available Tools

> 📚 **For detailed instructions, constraints, and JSON examples for all tools, please see the [Tool Guide](TOOL_GUIDE.md).**

| Tool | Description |
|------|-------------|
| `excel_create` | Create Excel workbooks with sheets, charts, formulas, and styling |
| `excel_export` | Export data to xlsx, ods, csv formats |
| `docx_create` | Create structured Word documents with sections (title, subtitle, TOC, headings, lists, tables) |
| `docx_export` | Export documents to docx, odt formats |
| `pptx_create` | Create PowerPoint presentations with slides, tables, and corporate branding |
| `pptx_export` | Export presentations to pptx, odp formats |
| `list_themes_tool` | List all available themes |
| `list_files` | List generated files for a session |
| `get_storage_stats` | Get storage usage statistics |

---

## Usage Examples

### Create an Excel Workbook

Supports formulas (cells starting with `=`), charts, and multi-sheet layouts.

```json
{
  "tool": "excel_create",
  "arguments": {
    "filename": "sales_report.xlsx",
    "sheets": [
      {
        "name": "Revenue",
        "headers": ["Month", "Sales", "Expenses", "Profit"],
        "rows": [
          ["Jan", 50000, 30000, "=B2-C2"],
          ["Feb", 62000, 35000, "=B3-C3"],
          ["Total", "=SUM(B2:B3)", "=SUM(C2:C3)", "=SUM(D2:D3)"]
        ],
        "charts": [
          {"chart_type": "bar", "data_range": "A1:B3", "title": "Monthly Sales", "position": "F2"}
        ]
      }
    ],
    "theme": "corporate"
  }
}
```

### Create a Word Document

Uses a structured `sections` array to build rich documents with proper heading hierarchy, table of contents, lists, and interleaved tables.

```json
{
  "tool": "docx_create",
  "arguments": {
    "filename": "quarterly_report.docx",
    "theme": "corporate",
    "sections": [
      {"type": "title", "text": "Quarterly Report"},
      {"type": "subtitle", "text": "Q3 2026 Financial Summary"},
      {"type": "toc"},
      {"type": "heading_1", "text": "1. Revenue Analysis"},
      {"type": "paragraph", "text": "Revenue grew by 12% year-over-year."},
      {"type": "list_bullet", "items": ["SaaS: +15%", "Services: +8%"]},
      {"type": "heading_2", "text": "1.1 Regional Breakdown"},
      {"type": "list_number", "items": ["APAC grew fastest", "EMEA remained stable"]},
      {"type": "table", "headers": ["Region", "Revenue"], "rows": [["APAC", "$2.1M"], ["EMEA", "$1.8M"]]},
      {"type": "heading_1", "text": "2. Outlook"},
      {"type": "paragraph", "text": "We expect continued growth in Q4."}
    ]
  }
}
```

**Supported section types:**

| Type | Fields | Description |
|------|--------|-------------|
| `title` | `text` | Document title (Title style) |
| `subtitle` | `text` | Document subtitle (Subtitle style) |
| `toc` | — | Table of Contents (auto-populates from headings) |
| `heading_1` | `text` | Heading level 1 |
| `heading_2` | `text` | Heading level 2 |
| `heading_3` | `text` | Heading level 3 |
| `paragraph` | `text` | Body paragraph |
| `list_bullet` | `items` | Bulleted list |
| `list_number` | `items` | Numbered list |
| `table` | `headers`, `rows` | Data table with themed styling |

> **Note:** The Table of Contents uses Word field codes. Right-click the TOC and select "Update Field" in Word/WPS Office to populate it.

### Create a PowerPoint Presentation

Pass a `slides` array with explicit slide objects for precise control over every slide.

```json
{
  "tool": "pptx_create",
  "arguments": {
    "filename": "project_update.pptx",
    "theme": "corporate",
    "slides": [
      {"title": "Welcome", "layout": "title"},
      {"title": "Key Metrics", "content": "All targets exceeded", "bullets": ["Revenue +12%", "Users +25%"], "layout": "title_and_content"},
      {"title": "Financial Summary", "layout": "title_and_content", "table_headers": ["Metric", "Value"], "table_rows": [["Revenue", "$2.1M"], ["Profit", "$400K"]]},
      {"title": "Next Steps", "content": "Questions?", "layout": "title_only"}
    ]
  }
}
```

**Available slide layouts:** `title`, `title_and_content`, `title_only`, `two_content`, `blank`, `section_header`

---

## Custom Templates

Use your own branded document as a base. The server preserves all formatting, styles, formulas, charts, master slides, and layouts from your template.

```json
{
  "tool": "docx_create",
  "arguments": {
    "filename": "report.docx",
    "template_path": "/path/to/your_template.docx",
    "sections": [
      {"type": "heading_1", "text": "Project Status"},
      {"type": "paragraph", "text": "All milestones on track."}
    ]
  }
}
```

Supported template formats: `.xlsx`, `.docx`, `.pptx`

---

## Themes

| Theme | Description | Primary Color |
|-------|-------------|---------------|
| `corporate` | Professional blue tones | #1E40AF |
| `minimal` | Clean black and white | #000000 |
| `creative` | Warm cream with amber tones | #78350F |
| `academic` | Formal with serif fonts | #1E3A5F |
| `dark` | Dark mode with slate tones | #1E293B |

---

## Configuration

| Environment Variable | Default | Description |
|---------------------|---------|-------------|
| `S3_ENDPOINT`       | `None`  | Ceph RADOS S3 endpoint (e.g. https://dcs3.psn.co.id) |
| `S3_BUCKET_NAME`    | `None`  | S3 target bucket |
| `S3_REGION`         | `None`  | S3 region |
| `S3_ACCESS_KEY`     | `None`  | S3 access key |
| `S3_SECRET_KEY`     | `None`  | S3 secret key |
| `OUTPUT_DIR` | `outputs` | Temporary local cache directory |
| `FILE_RETENTION_HOURS` | `24` | Hours before local gzip cache is deleted |
| `RATE_LIMIT_REQUESTS` | `20` | Max requests per window per user |
| `RATE_LIMIT_WINDOW` | `60` | Rate limit window in seconds |
| `MCP_TRANSPORT` | `stdio` | Transport mode (stdio or sse) |
| `LOCALE` | `en_US` | Default locale (en_US, id_ID) |

---

## Minimum Specifications

To run the MCP Office Server with all Data Analysis features (Track B), the following minimum specifications are recommended:

- **CPU:** 2 Cores (4 Cores recommended for heavy batch generation)
- **RAM:** 2 GB Minimum (4 GB+ recommended if processing large datasets >100,000 rows with pandas)
- **Storage:** 500 MB free space (for Docker images, dependencies like `numpy`/`scipy`/`pandas`, and local output caching)
- **Python:** Version 3.11 or higher
- **OS:** Windows, macOS, or Linux

---

## Project Structure

```
mcp-office/
├── src/
│   ├── server.py                # MCP server entry point
│   ├── generators/
│   │   ├── excel_generator.py   # Excel (.xlsx)
│   │   ├── docx_generator.py    # Word (.docx)
│   │   ├── pptx_generator.py    # PowerPoint (.pptx)
│   │   └── odf_generator.py     # ODF (.ods, .odt, .odp)
│   ├── tools/
│   │   ├── definitions.py       # Tool schemas and JSON guides
│   │   └── handlers.py          # Tool handler implementations
│   ├── styles/
│   │   ├── themes.py            # Theme definitions
│   │   └── style_applier.py     # Style application utilities
│   ├── templates/               # Base document templates
│   └── utils/
│       ├── file_handler.py      # File operations
│       ├── cleanup.py           # Auto cleanup
│       ├── rate_limiter.py      # Rate limiting
│       ├── validators.py        # Input validation
│       ├── security.py          # PII redaction, sanitization
│       └── logger.py            # Structured logging
├── run_server.py                # Server launcher
├── Dockerfile
├── docker-compose.yml
├── pyproject.toml
└── README.md
```

## Dependencies

| Package | Purpose |
|---------|---------|
| **mcp** | MCP server framework |
| **openpyxl** | Excel file generation |
| **python-docx** | Word document generation |
| **python-pptx** | PowerPoint generation |
| **odfpy** | ODF format support |
| **pydantic** | Data validation |
| **pydantic-settings** | Environment configuration |
| **loguru** | Structured logging |
| **aiofiles** | Async file I/O |
| **aiohttp** | HTTP client |
| **uvicorn** | ASGI server (SSE mode) |
| **starlette** | ASGI framework (SSE mode) |

## License

MIT