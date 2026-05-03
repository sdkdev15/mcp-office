"""MCP Office Server — Main entry point for mcp 1.27.x."""

from __future__ import annotations

import asyncio
import os
import uuid
from typing import Any, Optional

from mcp.server import Server
from mcp.server.models import InitializationOptions
from mcp.types import TextContent, Tool, ServerCapabilities
from mcp.server.stdio import stdio_server
from mcp.server.sse import SseServerTransport

from src.generators.excel_generator import ExcelGenerator
from src.generators.docx_generator import DOCXGenerator
from src.generators.pptx_generator import PPTXGenerator
from src.generators.odf_generator import ODFGenerator
from src.styles.themes import list_themes
from src.utils.file_handler import FileHandler
from src.utils.cleanup import FileCleanup
from src.utils.rate_limiter import RateLimiter
from src.utils.logger import get_logger
from src.utils.validators import ValidationError

# ── Initialize Components ──

log = get_logger("server")
app = Server("mcp-office", version="1.0.0")
file_handler = FileHandler(output_dir=os.environ.get("OUTPUT_DIR", "outputs"))
file_cleanup = FileCleanup(output_dir=os.environ.get("OUTPUT_DIR", "outputs"))
rate_limiter = RateLimiter()


def get_session_id(params: dict[str, Any]) -> str:
    return params.get("session_id") or str(uuid.uuid4())[:8]


def check_rate_limit(session_id: str) -> Optional[str]:
    allowed, info = rate_limiter.is_allowed(session_id)
    if not allowed:
        return f"Rate limit exceeded. Try again in {info.get('retry_after', 60)}s."
    return None


def format_file_result(file_info: dict) -> str:
    return (
        f"✅ **File Generated Successfully**\n\n"
        f"**File:** {file_info['filename']}\n"
        f"**Size:** {file_handler._human_readable_size(file_info['size_bytes'])}\n"
        f"**Type:** {file_info['mime_type']}\n"
        f"**Path:** {file_info['filepath']}\n"
        f"**Created:** {file_info['created_at']}"
    )


# ── Tool Definitions ──

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
                    "description": "List of sheet objects with 'name', 'headers', and 'rows'",
                    "items": {"type": "object"},
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
                "sheets": {"type": "array", "items": {"type": "object"}},
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
                "tables": {"type": "array", "description": "Optional list of table data with headers and rows", "items": {"type": "object"}},
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
        description="Create a PowerPoint presentation (.pptx) with slides and theme styling.",
        inputSchema={
            "type": "object",
            "properties": {
                "filename": {"type": "string"},
                "title": {"type": "string", "default": "Presentation"},
                "theme": {"type": "string", "default": "corporate"},
                "template_path": {"type": "string", "description": "Optional path to a .pptx template file to use as base (preserves master slides, themes, layouts)"},
                "slides": {"type": "array", "description": "Optional list of slide data with title, content, bullets", "items": {"type": "object"}},
                "slide_size": {"type": "string", "default": "widescreen"},
                "session_id": {"type": "string"},
                "metadata": {"type": "object"},
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


# ── MCP Handlers ──

@app.list_tools()
async def list_tools() -> list[Tool]:
    """List available tools."""
    return TOOLS


@app.call_tool()
async def call_tool(name: str, arguments: dict[str, Any]) -> list[TextContent]:
    """Handle tool calls."""
    session_id = get_session_id(arguments)

    if name not in ("list_themes_tool", "list_files", "get_storage_stats"):
        rate_error = check_rate_limit(session_id)
        if rate_error:
            return [TextContent(type="text", text=rate_error)]

    try:
        if name == "excel_create":
            return await _excel_create(arguments)
        elif name == "excel_export":
            return await _excel_export(arguments)
        elif name == "docx_create":
            return await _docx_create(arguments)
        elif name == "docx_generate_from_prompt":
            return await _docx_generate_from_prompt(arguments)
        elif name == "docx_export":
            return await _docx_export(arguments)
        elif name == "pptx_create":
            return await _pptx_create(arguments)
        elif name == "pptx_generate_from_prompt":
            return await _pptx_generate_from_prompt(arguments)
        elif name == "pptx_export":
            return await _pptx_export(arguments)
        elif name == "list_themes_tool":
            return await _list_themes()
        elif name == "list_files":
            return await _list_files(arguments)
        elif name == "get_storage_stats":
            return await _get_storage_stats()
        else:
            return [TextContent(type="text", text=f"Unknown tool: {name}")]

    except ValidationError as e:
        return [TextContent(type="text", text=f"Validation Error: {e.message}")]
    except Exception as e:
        log.error(f"Tool {name} failed: {e}")
        return [TextContent(type="text", text=f"Error: {str(e)}")]


# ── Tool Implementations ──

async def _excel_create(args: dict) -> list[TextContent]:
    gen = ExcelGenerator(args.get("theme", "corporate"))
    template_path = args.get("template_path")
    
    if template_path:
        # Use user-provided template as base
        data = gen.create_from_template(
            template_path,
            args["sheets"],
            args.get("metadata"),
        )
    else:
        data = gen.create_workbook(args["sheets"], args.get("metadata"))
    
    filename = args["filename"]
    if not filename.endswith(".xlsx"):
        filename += ".xlsx"
    filename = file_handler.generate_filename(filename.rsplit(".", 1)[0], ".xlsx")
    file_info = await file_handler.save_file(data, filename, get_session_id(args))
    return [TextContent(type="text", text=format_file_result(file_info))]


async def _excel_export(args: dict) -> list[TextContent]:
    results = []
    fmt = args.get("format", "all").lower()
    session_id = get_session_id(args)

    if fmt in ("xlsx", "all"):
        gen = ExcelGenerator()
        data = gen.create_workbook(args["sheets"])
        filename = file_handler.generate_filename("export", ".xlsx")
        file_info = await file_handler.save_file(data, filename, session_id)
        results.append(format_file_result(file_info))

    if fmt in ("ods", "all"):
        gen = ODFGenerator()
        data = gen.create_spreadsheet(args["sheets"])
        filename = file_handler.generate_filename("export", ".ods")
        file_info = await file_handler.save_file(data, filename, session_id)
        results.append(format_file_result(file_info))

    return [TextContent(type="text", text="\n\n".join(results))]


async def _docx_create(args: dict) -> list[TextContent]:
    gen = DOCXGenerator(args.get("theme", "corporate"))
    template_path = args.get("template_path")
    
    if template_path:
        # Use user-provided template as base
        data = gen.create_from_template(
            template_path,
            title=args.get("title", "Document"),
            content_paragraphs=args.get("content_paragraphs"),
            tables=args.get("tables"),
            metadata=args.get("metadata"),
        )
    else:
        data = gen.create_document(
            args.get("title", "Document"),
            args.get("page_size", "A4"),
            args.get("orientation", "portrait"),
            args.get("metadata"),
        )
    
    filename = args["filename"]
    if not filename.endswith(".docx"):
        filename += ".docx"
    filename = file_handler.generate_filename(filename.rsplit(".", 1)[0], ".docx")
    file_info = await file_handler.save_file(data, filename, get_session_id(args))
    return [TextContent(type="text", text=format_file_result(file_info))]


async def _docx_generate_from_prompt(args: dict) -> list[TextContent]:
    gen = DOCXGenerator(args.get("theme", "corporate"))
    title = args.get("filename") or "Generated Document"
    data = gen.create_from_prompt(args["prompt"], title)
    filename = args.get("filename") or "document.docx"
    if not filename.endswith(".docx"):
        filename += ".docx"
    filename = file_handler.generate_filename(filename.rsplit(".", 1)[0], ".docx")
    file_info = await file_handler.save_file(data, filename, get_session_id(args))
    return [TextContent(type="text", text=format_file_result(file_info))]


async def _docx_export(args: dict) -> list[TextContent]:
    results = []
    fmt = args.get("format", "all").lower()
    session_id = get_session_id(args)
    title = args.get("title", "Document")

    if fmt in ("docx", "all"):
        gen = DOCXGenerator(args.get("theme", "corporate"))
        data = gen.create_document(title, metadata=args.get("metadata"))
        filename = file_handler.generate_filename(title, ".docx")
        file_info = await file_handler.save_file(data, filename, session_id)
        results.append(format_file_result(file_info))

    if fmt in ("odt", "all"):
        gen = ODFGenerator()
        odt_doc = gen.create_text_document(title, args.get("metadata"))
        gen.add_paragraph_to_odt(odt_doc, f"Generated: {title}")
        data = gen.save_odt(odt_doc)
        filename = file_handler.generate_filename(title, ".odt")
        file_info = await file_handler.save_file(data, filename, session_id)
        results.append(format_file_result(file_info))

    return [TextContent(type="text", text="\n\n".join(results))]


async def _pptx_create(args: dict) -> list[TextContent]:
    gen = PPTXGenerator(args.get("theme", "corporate"))
    template_path = args.get("template_path")
    
    if template_path:
        # Use user-provided template as base
        data = gen.create_from_template(
            template_path,
            slides=args.get("slides"),
            metadata=args.get("metadata"),
        )
    else:
        data = gen.create_presentation(
            args.get("title", "Presentation"),
            args.get("slide_size", "widescreen"),
            args.get("metadata"),
        )
    
    filename = args["filename"]
    if not filename.endswith(".pptx"):
        filename += ".pptx"
    filename = file_handler.generate_filename(filename.rsplit(".", 1)[0], ".pptx")
    file_info = await file_handler.save_file(data, filename, get_session_id(args))
    return [TextContent(type="text", text=format_file_result(file_info))]


async def _pptx_generate_from_prompt(args: dict) -> list[TextContent]:
    gen = PPTXGenerator(args.get("theme", "corporate"))
    title = args.get("filename") or "Generated Presentation"
    data = gen.create_from_prompt(args["prompt"], title)
    filename = args.get("filename") or "presentation.pptx"
    if not filename.endswith(".pptx"):
        filename += ".pptx"
    filename = file_handler.generate_filename(filename.rsplit(".", 1)[0], ".pptx")
    file_info = await file_handler.save_file(data, filename, get_session_id(args))
    return [TextContent(type="text", text=format_file_result(file_info))]


async def _pptx_export(args: dict) -> list[TextContent]:
    results = []
    fmt = args.get("format", "all").lower()
    session_id = get_session_id(args)
    title = args.get("title", "Presentation")

    if fmt in ("pptx", "all"):
        gen = PPTXGenerator(args.get("theme", "corporate"))
        data = gen.create_presentation(title, metadata=args.get("metadata"))
        filename = file_handler.generate_filename(title, ".pptx")
        file_info = await file_handler.save_file(data, filename, session_id)
        results.append(format_file_result(file_info))

    if fmt in ("odp", "all"):
        gen = ODFGenerator()
        odp_doc = gen.create_presentation(title, args.get("metadata"))
        data = gen.save_odp(odp_doc)
        filename = file_handler.generate_filename(title, ".odp")
        file_info = await file_handler.save_file(data, filename, session_id)
        results.append(format_file_result(file_info))

    return [TextContent(type="text", text="\n\n".join(results))]


async def _list_themes() -> list[TextContent]:
    themes = list_themes()
    result = "**Available Themes:**\n\n"
    for theme in themes:
        result += f"- **{theme['name']}**: {theme['description']}\n  - Primary: {theme['primary_color']}\n\n"
    return [TextContent(type="text", text=result)]


async def _list_files(args: dict) -> list[TextContent]:
    session_id = args.get("session_id", "")
    files = file_handler.list_session_files(session_id)
    if not files:
        return [TextContent(type="text", text=f"No files found for session: {session_id}")]
    result = f"**Files in session {session_id}:**\n\n"
    for f in files:
        result += f"- `{f['filename']}` ({f['size_bytes']} bytes)\n  - Created: {f['created_at']}\n\n"
    return [TextContent(type="text", text=result)]


async def _get_storage_stats() -> list[TextContent]:
    stats = file_cleanup.get_storage_stats()
    result = (
        f"**Storage Statistics:**\n\n"
        f"- Total files: {stats['total_files']}\n"
        f"- Total size: {stats['total_size_human']}\n"
        f"- Active sessions: {stats['active_sessions']}\n"
        f"- Retention: {stats['retention_hours']} hours"
    )
    return [TextContent(type="text", text=result)]


# ── Server Startup ──

async def run_stdio_server():
    """Run the MCP server using stdio transport."""
    log.info("Starting MCP Office Server (stdio transport)")
    file_cleanup.start()

    async with stdio_server() as (read_stream, write_stream):
        await app.run(
            read_stream,
            write_stream,
            InitializationOptions(
                server_name="mcp-office",
                server_version="1.0.0",
                capabilities=ServerCapabilities(),
            ),
        )

    file_cleanup.stop()
    log.info("MCP Office Server stopped")


async def run_sse_server():
    """Run the MCP server using SSE transport (for remote/server deployment)."""
    import uvicorn
    from starlette.applications import Starlette
    from starlette.routing import Route, Mount
    from starlette.requests import Request

    log.info("Starting MCP Office Server (SSE transport on port 8765)")
    file_cleanup.start()

    sse_transport = SseServerTransport("/messages/")

    # ASGI app for SSE — handles its own HTTP response lifecycle
    async def sse_asgi(scope, receive, send):
        if scope["type"] != "http":
            return
        log.info("=== New SSE connection ===")
        try:
            async with sse_transport.connect_sse(
                scope, receive, send
            ) as (read_stream, write_stream):
                log.info("SSE streams established, starting app.run()")
                await app.run(
                    read_stream,
                    write_stream,
                    InitializationOptions(
                        server_name="mcp-office",
                        server_version="1.0.0",
                        capabilities=ServerCapabilities(),
                    ),
                )
                log.info("app.run() completed normally")
        except Exception as e:
            log.error(f"SSE connection error: {e}")

    # ASGI app for messages — handle_post_message manages its own response via send()
    async def messages_asgi(scope, receive, send):
        if scope["type"] != "http":
            return
        if scope["method"] != "POST":
            from starlette.responses import PlainTextResponse
            response = PlainTextResponse("Method Not Allowed", status_code=405)
            await response(scope, receive, send)
            return
        log.info("Message POST received")
        await sse_transport.handle_post_message(scope, receive, send)

    starlette_app = Starlette(
        debug=True,
        routes=[
            Mount("/sse", app=sse_asgi),
            Mount("/messages/", app=messages_asgi),
        ],
    )

    config = uvicorn.Config(starlette_app, host="0.0.0.0", port=8765, log_level="info")
    server = uvicorn.Server(config)
    await server.serve()

    file_cleanup.stop()
    log.info("MCP Office Server (SSE) stopped")


async def main():
    """Main entry point — supports both stdio and SSE transports."""
    transport = os.environ.get("MCP_TRANSPORT", "stdio").lower()
    if transport == "sse":
        await run_sse_server()
    else:
        await run_stdio_server()


if __name__ == "__main__":
    asyncio.run(main())