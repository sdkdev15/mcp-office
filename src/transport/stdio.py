"""Stdio transport for the MCP Office server."""

from __future__ import annotations

from mcp.server import Server
from mcp.server.models import InitializationOptions
from mcp.types import ServerCapabilities
from mcp.server.stdio import stdio_server

from src.utils.cleanup import FileCleanup
from src.utils.logger import get_logger

log = get_logger("stdio")


async def run_stdio_server(app: Server, file_cleanup: FileCleanup) -> None:
    """Run the MCP server using stdio transport.

    Args:
        app: Configured MCP Server instance.
        file_cleanup: FileCleanup instance for background cleanup.
    """
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
