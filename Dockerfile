# [final]
FROM python:3.11-slim

WORKDIR /app

# Install runtime dependencies directly
COPY pyproject.toml ./
RUN pip install --no-cache-dir \
    openpyxl \
    python-docx \
    python-pptx \
    odfpy \
    mcp \
    loguru \
    pydantic \
    pydantic-settings \
    aiofiles \
    aiohttp \
    uvicorn \
    starlette

# Create non-root user
RUN useradd -m -u 1000 appuser && \
    mkdir -p /app/outputs && \
    chown -R appuser:appuser /app

USER appuser

# Copy application code
COPY --chown=appuser:appuser src/ ./src/
COPY --chown=appuser:appuser pyproject.toml ./

# Environment variables
ENV PYTHONUNBUFFERED=1 \
    PYTHONDONTWRITEBYTECODE=1 \
    OUTPUT_DIR=/app/outputs \
    FILE_RETENTION_HOURS=24 \
    RATE_LIMIT_REQUESTS=20 \
    RATE_LIMIT_WINDOW=60 \
    MCP_TRANSPORT=stdio \
    LOCALE=en_US

# Health check
HEALTHCHECK --interval=30s --timeout=5s --start-period=10s --retries=3 \
    CMD python -c "import src.server; print('healthy')" || exit 1

# Default command (stdio mode)
CMD ["python", "-m", "src.server"]
