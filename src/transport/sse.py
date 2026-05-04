"""SSE transport for the MCP Office server.

Includes the ASGI router for SSE connections, message handling,
and file downloads with path traversal protection.
"""

from __future__ import annotations

import os
import pathlib

from mcp.server import Server
from mcp.server.models import InitializationOptions
from mcp.types import ServerCapabilities
from mcp.server.sse import SseServerTransport

from src.utils.cleanup import FileCleanup
from src.utils.logger import get_logger

log = get_logger("sse")


async def run_sse_server(app: Server, file_cleanup: FileCleanup) -> None:
    """Run the MCP server using SSE transport (for remote/server deployment).

    Args:
        app: Configured MCP Server instance.
        file_cleanup: FileCleanup instance for background cleanup.
    """
    import uvicorn
    from starlette.applications import Starlette
    from starlette.routing import Mount

    log.info("Starting MCP Office Server (SSE transport on port 8765)")
    file_cleanup.start()

    sse_transport = SseServerTransport("/messages/")
    output_dir = pathlib.Path(os.environ.get("OUTPUT_DIR", "outputs")).resolve()

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

    # ASGI app for messages
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

    # File download handler with path traversal protection (#17)
    async def files_asgi(scope, receive, send):
        if scope["type"] != "http":
            return
        from starlette.responses import FileResponse, PlainTextResponse, StreamingResponse
        import boto3
        from botocore.config import Config

        path = scope.get("path", "")
        parts = path.split("/")
        if len(parts) < 4 or parts[1] != "files":
            response = PlainTextResponse("Not Found", status_code=404)
            await response(scope, receive, send)
            return

        session_id = parts[2]
        filename = "/".join(parts[3:])
        s3_key = f"{session_id}/{filename}"
        
        endpoint = os.environ.get("S3_ENDPOINT")
        bucket = os.environ.get("S3_BUCKET_NAME")
        
        if endpoint and bucket:
            region = os.environ.get("S3_REGION")
            access_key = os.environ.get("S3_ACCESS_KEY")
            secret_key = os.environ.get("S3_SECRET_KEY")
            
            s3 = boto3.client(
                "s3",
                endpoint_url=endpoint,
                region_name=region,
                aws_access_key_id=access_key,
                aws_secret_access_key=secret_key,
                config=Config(s3={'addressing_style': 'path'}, signature_version='s3v4')
            )
            try:
                s3_resp = s3.get_object(Bucket=bucket, Key=s3_key)
                def generate():
                    for chunk in s3_resp['Body'].iter_chunks(chunk_size=8192):
                        yield chunk
                        
                content_type = s3_resp.get('ContentType', "application/octet-stream")
                response = StreamingResponse(generate(), media_type=content_type)
                response.headers["Content-Disposition"] = f'attachment; filename="{filename}"'
                await response(scope, receive, send)
                return
            except Exception as e:
                log.error(f"S3 download failed for {s3_key}, falling back to local backup: {e}")

        # Fallback to local .gz backup if S3 failed/not configured
        file_path = (output_dir / session_id / f"{filename}.gz").resolve()

        if not str(file_path).startswith(str(output_dir)):
            log.warning(f"Path traversal attempt blocked: {path}")
            response = PlainTextResponse("Forbidden", status_code=403)
            await response(scope, receive, send)
            return

        if not file_path.is_file():
            response = PlainTextResponse(f"File not found on S3 and local backup missing: {filename}", status_code=404)
            await response(scope, receive, send)
            return

        log.info(f"Serving local backup file uncompressed: {file_path}")
        
        import gzip
        import mimetypes
        
        def generate_local_gz():
            with gzip.open(file_path, 'rb') as f:
                while chunk := f.read(8192):
                    yield chunk

        mt, _ = mimetypes.guess_type(filename)
        mt = mt or "application/octet-stream"
        
        response = StreamingResponse(generate_local_gz(), media_type=mt)
        response.headers["Content-Disposition"] = f'attachment; filename="{filename}"'
        await response(scope, receive, send)

    # Single ASGI router — avoids Starlette Mount 307 redirect on /sse → /sse/
    async def router_asgi(scope, receive, send):
        if scope["type"] != "http":
            return
        path = scope.get("path", "")
        if path == "/sse" or path == "/sse/":
            await sse_asgi(scope, receive, send)
        elif path.startswith("/messages"):
            await messages_asgi(scope, receive, send)
        elif path.startswith("/files"):
            await files_asgi(scope, receive, send)
        else:
            from starlette.responses import PlainTextResponse
            response = PlainTextResponse("Not Found", status_code=404)
            await response(scope, receive, send)

    debug_mode = os.environ.get("DEBUG", "false").lower() == "true"
    starlette_app = Starlette(
        debug=debug_mode,
        routes=[
            Mount("/", app=router_asgi),
        ],
    )

    config = uvicorn.Config(starlette_app, host="0.0.0.0", port=8765, log_level="info")
    server = uvicorn.Server(config)
    await server.serve()

    file_cleanup.stop()
    log.info("MCP Office Server (SSE) stopped")
