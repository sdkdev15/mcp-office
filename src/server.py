"""MCP Office Server — Main entry point for mcp 1.27.x.

This module bootstraps the server, wires up components, and
dispatches to the appropriate transport (stdio or SSE).
"""

from __future__ import annotations

import asyncio
import os

from mcp.server import Server

from src.tools.handlers import register_handlers
from src.utils.file_handler import FileHandler
from src.utils.cleanup import FileCleanup
from src.utils.rate_limiter import RateLimiter
from src.utils.logger import get_logger

log = get_logger("server")


def create_app() -> tuple[Server, FileHandler, FileCleanup, RateLimiter]:
    """Create and configure the MCP server with all components.

    Returns:
        Tuple of (app, file_handler, file_cleanup, rate_limiter).
    """
    app = Server("mcp-office", version="1.0.0")
    file_handler = FileHandler(output_dir=os.environ.get("OUTPUT_DIR", "outputs"))
    file_cleanup = FileCleanup(output_dir=os.environ.get("OUTPUT_DIR", "outputs"))
    rate_limiter = RateLimiter()
    base_url = os.environ.get("BASE_URL", "")

    register_handlers(
        app,
        file_handler=file_handler,
        file_cleanup=file_cleanup,
        rate_limiter=rate_limiter,
        base_url=base_url,
    )

    return app, file_handler, file_cleanup, rate_limiter


async def main():
    """Main entry point — supports both stdio and SSE transports."""
    app, file_handler, file_cleanup, rate_limiter = create_app()

    transport = os.environ.get("MCP_TRANSPORT", "stdio").lower()
    if transport == "sse":
        from src.transport.sse import run_sse_server
        await run_sse_server(app, file_cleanup)
    else:
        from src.transport.stdio import run_stdio_server
        await run_stdio_server(app, file_cleanup)


if __name__ == "__main__":
    asyncio.run(main())