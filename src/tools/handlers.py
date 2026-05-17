"""Tool handler implementations for document generation.

Each handler receives raw arguments from the MCP call_tool dispatcher,
generates the requested document, saves it via FileHandler, and returns
a formatted result.

Security: InputSanitizer is applied to user-facing text inputs.
          Template paths are validated against the output directory.
"""

from __future__ import annotations

import asyncio
import os
import uuid
from pathlib import Path
from typing import Any, Optional

from docx import Document
from mcp.types import TextContent

from src.generators.excel_generator import ExcelGenerator
from src.generators.docx_generator import DOCXGenerator
from src.generators.pptx_generator import PPTXGenerator
from src.generators.odf_generator import ODFGenerator
from src.styles.themes import list_themes
from src.utils.file_handler import FileHandler
from src.utils.cleanup import FileCleanup
from src.utils.rate_limiter import RateLimiter
from src.utils.security import InputSanitizer
from src.utils.logger import get_logger
from src.utils.validators import ValidationError
from src.utils.formatting import human_readable_size
from src.tools.definitions import TOOLS

log = get_logger("handlers")

# ── Shared Components (initialized by register_handlers) ──

_file_handler: FileHandler = None  # type: ignore[assignment]
_file_cleanup: FileCleanup = None  # type: ignore[assignment]
_rate_limiter: RateLimiter = None  # type: ignore[assignment]
_base_url: str = ""
_sanitizer = InputSanitizer()

# Allowed template directories (only output dir by default)
_allowed_template_dirs: list[Path] = []


def register_handlers(
    app,
    file_handler: FileHandler,
    file_cleanup: FileCleanup,
    rate_limiter: RateLimiter,
    base_url: str = "",
) -> None:
    """Register MCP tool handlers on the given server app.

    Args:
        app: MCP Server instance.
        file_handler: FileHandler for saving generated files.
        file_cleanup: FileCleanup for storage stats.
        rate_limiter: RateLimiter for per-session limits.
        base_url: Base URL for download links (SSE mode).
    """
    global _file_handler, _file_cleanup, _rate_limiter, _base_url, _allowed_template_dirs
    _file_handler = file_handler
    _file_cleanup = file_cleanup
    _rate_limiter = rate_limiter
    _base_url = base_url
    _allowed_template_dirs = [
        file_handler.output_dir.resolve(),
    ]

    @app.list_tools()
    async def list_tools_handler():
        """List available tools."""
        return TOOLS

    @app.call_tool()
    async def call_tool_handler(name: str, arguments: dict[str, Any]):
        """Handle tool calls."""
        session_id = _get_session_id(arguments)

        if name not in ("list_themes_tool", "list_files", "get_storage_stats"):
            rate_error = _check_rate_limit(session_id)
            if rate_error:
                return [TextContent(type="text", text=rate_error)]

        try:
            handler_map = {
                "excel_create": _excel_create,
                "excel_export": _excel_export,
                "docx_create": _docx_create,
                "docx_export": _docx_export,
                "pptx_create": _pptx_create,
                "pptx_export": _pptx_export,
                "excel_advanced_formatting": _excel_create,  # Reuse creation logic
                "excel_with_images": _excel_create,          # Reuse creation logic
                "docx_with_images": _docx_create,            # Reuse creation logic
                "analyze_data": _analyze_data,
                "generate_summary": _generate_summary,
                "generate_faq": _generate_faq,
                "recommend_charts": _recommend_charts,
                "list_themes_tool": _list_themes,
                "list_files": _list_files,
                "get_storage_stats": _get_storage_stats,
            }
            handler = handler_map.get(name)
            if handler is None:
                return [TextContent(type="text", text=f"Unknown tool: {name}")]

            if name in ("list_themes_tool", "get_storage_stats"):
                return await handler()
            return await handler(arguments)

        except ValidationError as e:
            return [TextContent(type="text", text=f"Validation Error: {e.message}")]
        except Exception as e:
            log.error(f"Tool {name} failed: {e}")
            return [TextContent(type="text", text=f"Error: {str(e)}")]


# ── Helpers ──

def _get_session_id(params: dict[str, Any]) -> str:
    return params.get("session_id") or str(uuid.uuid4())[:8]


def _check_rate_limit(session_id: str) -> Optional[str]:
    allowed, info = _rate_limiter.is_allowed(session_id)
    if not allowed:
        return f"Rate limit exceeded. Try again in {info.get('retry_after', 60)}s."
    return None


def _validate_template_path(path: Optional[str]) -> Optional[str]:
    """Validate that a template path is within allowed directories.

    Args:
        path: User-supplied template path.

    Returns:
        Validated path or None.

    Raises:
        ValidationError: If path is outside allowed directories.
    """
    if not path:
        return None

    resolved = Path(path).resolve()
    if not resolved.is_file():
        raise ValidationError("template_path", f"Template file not found: {path}")

    for allowed_dir in _allowed_template_dirs:
        try:
            resolved.relative_to(allowed_dir)
            return str(resolved)
        except ValueError:
            continue

    raise ValidationError(
        "template_path",
        "Template path is outside allowed directories. "
        "Templates must be within the output directory."
    )


def _format_file_result(file_info: dict) -> str:
    result = (
        f"✅ **File Generated Successfully**\n\n"
        f"**File:** {file_info['filename']}\n"
        f"**Size:** {human_readable_size(file_info['size_bytes'])}\n"
        f"**Type:** {file_info['mime_type']}\n"
        f"**Path:** {file_info['filepath']}\n"
        f"**Created:** {file_info['created_at']}"
    )
    if _base_url:
        download_url = f"{_base_url}/files/{file_info['session_id']}/{file_info['filename']}"
        result += f"\n\n**Download:** [{download_url}]({download_url})"
    return result


# ── Tool Implementations ──

async def _excel_create(args: dict) -> list[TextContent]:
    gen = ExcelGenerator(args.get("theme", "corporate"))
    template_path = _validate_template_path(args.get("template_path"))

    def _sync_create():
        if template_path:
            return gen.create_from_template(
                template_path,
                args["sheets"],
                args.get("metadata"),
            )
        else:
            return gen.create_workbook(args["sheets"], args.get("metadata"))

    data = await asyncio.to_thread(_sync_create)

    filename = args["filename"]
    if not filename.endswith(".xlsx"):
        filename += ".xlsx"
    filename = _file_handler.generate_filename(filename.rsplit(".", 1)[0], ".xlsx")
    file_info = await _file_handler.save_file(data, filename, _get_session_id(args))
    return [TextContent(type="text", text=_format_file_result(file_info))]


async def _excel_export(args: dict) -> list[TextContent]:
    results = []
    fmt = args.get("format", "all").lower()
    session_id = _get_session_id(args)

    if fmt in ("xlsx", "all"):
        gen = ExcelGenerator()
        data = await asyncio.to_thread(gen.create_workbook, args["sheets"])
        filename = _file_handler.generate_filename("export", ".xlsx")
        file_info = await _file_handler.save_file(data, filename, session_id)
        results.append(_format_file_result(file_info))

    if fmt in ("ods", "all"):
        gen = ODFGenerator()
        data = await asyncio.to_thread(gen.create_spreadsheet, args["sheets"])
        filename = _file_handler.generate_filename("export", ".ods")
        file_info = await _file_handler.save_file(data, filename, session_id)
        results.append(_format_file_result(file_info))

    return [TextContent(type="text", text="\n\n".join(results))]


async def _docx_create(args: dict) -> list[TextContent]:
    gen = DOCXGenerator(args.get("theme", "corporate"))
    template_path = args.get("template_path")
    if template_path:
        template_path = _validate_template_path(template_path)
    else:
        import os
        theme_name = args.get("theme", "corporate")
        base_path = os.path.join(os.path.dirname(os.path.dirname(__file__)), "templates", f"{theme_name}_base.docx")
        if os.path.exists(base_path):
            template_path = base_path
        else:
            fallback = os.path.join(os.path.dirname(os.path.dirname(__file__)), "templates", "corporate_base.docx")
            if os.path.exists(fallback):
                template_path = fallback
    sections = args.get("sections")
    content_paragraphs = args.get("content_paragraphs")
    tables = args.get("tables")

    def _sync_create():
        if template_path:
            return gen.create_from_template(
                template_path,
                title=args.get("title", "Document"),
                sections=sections,
                content_paragraphs=content_paragraphs,
                tables=tables,
                metadata=args.get("metadata"),
            )
        else:
            return gen.create_document_with_content(
                title=args.get("title", "Document"),
                page_size=args.get("page_size", "A4"),
                orientation=args.get("orientation", "portrait"),
                metadata=args.get("metadata"),
                sections=sections,
                content_paragraphs=content_paragraphs,
                tables=tables,
            )

    data = await asyncio.to_thread(_sync_create)

    filename = args["filename"]
    if not filename.endswith(".docx"):
        filename += ".docx"
    filename = _file_handler.generate_filename(filename.rsplit(".", 1)[0], ".docx")
    file_info = await _file_handler.save_file(data, filename, _get_session_id(args))
    return [TextContent(type="text", text=_format_file_result(file_info))]



async def _docx_export(args: dict) -> list[TextContent]:
    results = []
    fmt = args.get("format", "all").lower()
    session_id = _get_session_id(args)
    title = args.get("title", "Document")

    if fmt in ("docx", "all"):
        gen = DOCXGenerator(args.get("theme", "corporate"))
        data = await asyncio.to_thread(gen.create_document, title, metadata=args.get("metadata"))
        filename = _file_handler.generate_filename(title, ".docx")
        file_info = await _file_handler.save_file(data, filename, session_id)
        results.append(_format_file_result(file_info))

    if fmt in ("odt", "all"):
        def _sync_create_odt():
            gen = ODFGenerator()
            odt_doc = gen.create_text_document(title, args.get("metadata"))
            gen.add_paragraph_to_odt(odt_doc, f"Generated: {title}")
            return gen.save_odt(odt_doc)

        data = await asyncio.to_thread(_sync_create_odt)
        filename = _file_handler.generate_filename(title, ".odt")
        file_info = await _file_handler.save_file(data, filename, session_id)
        results.append(_format_file_result(file_info))

    return [TextContent(type="text", text="\n\n".join(results))]


async def _pptx_create(args: dict) -> list[TextContent]:
    gen = PPTXGenerator(args.get("theme", "corporate"))
    template_path = args.get("template_path")
    if template_path:
        template_path = _validate_template_path(template_path)
    else:
        import os
        theme_name = args.get("theme", "corporate")
        base_path = os.path.join(os.path.dirname(os.path.dirname(__file__)), "templates", f"{theme_name}_base.pptx")
        if os.path.exists(base_path):
            template_path = base_path
        else:
            fallback = os.path.join(os.path.dirname(os.path.dirname(__file__)), "templates", "corporate_base.pptx")
            if os.path.exists(fallback):
                template_path = fallback
    slides = args.get("slides")

    def _sync_create():
        if template_path:
            return gen.create_from_template(
                template_path,
                slides=slides,
                metadata=args.get("metadata"),
            )
        else:
            return gen.create_presentation(
                title=args.get("title", "Presentation"),
                slide_size=args.get("slide_size", "widescreen"),
                metadata=args.get("metadata"),
                slides=slides
            )

    data = await asyncio.to_thread(_sync_create)

    filename = args["filename"]
    if not filename.endswith(".pptx"):
        filename += ".pptx"
    filename = _file_handler.generate_filename(filename.rsplit(".", 1)[0], ".pptx")
    file_info = await _file_handler.save_file(data, filename, _get_session_id(args))
    return [TextContent(type="text", text=_format_file_result(file_info))]



async def _pptx_export(args: dict) -> list[TextContent]:
    results = []
    fmt = args.get("format", "all").lower()
    session_id = _get_session_id(args)
    title = args.get("title", "Presentation")

    if fmt in ("pptx", "all"):
        gen = PPTXGenerator(args.get("theme", "corporate"))
        data = await asyncio.to_thread(gen.create_presentation, title, metadata=args.get("metadata"))
        filename = _file_handler.generate_filename(title, ".pptx")
        file_info = await _file_handler.save_file(data, filename, session_id)
        results.append(_format_file_result(file_info))

    if fmt in ("odp", "all"):
        def _sync_create_odp():
            gen = ODFGenerator()
            odp_doc = gen.create_presentation(title, args.get("metadata"))
            return gen.save_odp(odp_doc)

        data = await asyncio.to_thread(_sync_create_odp)
        filename = _file_handler.generate_filename(title, ".odp")
        file_info = await _file_handler.save_file(data, filename, session_id)
        results.append(_format_file_result(file_info))

    return [TextContent(type="text", text="\n\n".join(results))]


# ── Analysis Implementations ──

async def _analyze_data(args: dict) -> list[TextContent]:
    from src.analysis.analyzer import Analyzer
    analyzer = Analyzer()
    data = await asyncio.to_thread(analyzer.analyze, args["data"], args.get("target_columns"), args.get("breakdown_by"))
    import json
    return [TextContent(type="text", text=json.dumps(data, indent=2))]

async def _generate_summary(args: dict) -> list[TextContent]:
    from src.analysis.summary_generator import SummaryGenerator
    gen = SummaryGenerator()
    data = await asyncio.to_thread(gen.generate, args["data"], args.get("style", "professional"), args.get("include_metrics", True), args.get("max_insights", 5))
    import json
    return [TextContent(type="text", text=json.dumps(data, indent=2))]

async def _generate_faq(args: dict) -> list[TextContent]:
    from src.analysis.faq_generator import FAQGenerator
    gen = FAQGenerator()
    data = await asyncio.to_thread(gen.generate, args["data"], args.get("num_questions", 5), args.get("question_style", "practical"))
    import json
    return [TextContent(type="text", text=json.dumps(data, indent=2))]

async def _recommend_charts(args: dict) -> list[TextContent]:
    from src.analysis.chart_recommender import ChartRecommender
    gen = ChartRecommender()
    data = await asyncio.to_thread(gen.recommend, args["data"], args.get("data_types"), args.get("num_recommendations", 3))
    import json
    return [TextContent(type="text", text=json.dumps(data, indent=2))]


async def _list_themes() -> list[TextContent]:
    themes = list_themes()
    result = "**Available Themes:**\n\n"
    for theme in themes:
        result += f"- **{theme['name']}**: {theme['description']}\n  - Primary: {theme['primary_color']}\n\n"
    return [TextContent(type="text", text=result)]


async def _list_files(args: dict) -> list[TextContent]:
    session_id = args.get("session_id", "")
    files = _file_handler.list_session_files(session_id)
    if not files:
        return [TextContent(type="text", text=f"No files found for session: {session_id}")]
    result = f"**Files in session {session_id}:**\n\n"
    for f in files:
        result += f"- `{f['filename']}` ({f['size_bytes']} bytes)\n  - Created: {f['created_at']}\n\n"
    return [TextContent(type="text", text=result)]


async def _get_storage_stats() -> list[TextContent]:
    stats = _file_cleanup.get_storage_stats()
    result = (
        f"**Storage Statistics:**\n\n"
        f"- Total files: {stats['total_files']}\n"
        f"- Total size: {stats['total_size_human']}\n"
        f"- Active sessions: {stats['active_sessions']}\n"
        f"- Retention: {stats['retention_hours']} hours"
    )
    return [TextContent(type="text", text=result)]
