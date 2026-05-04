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
    lxml \
    mcp \
    loguru \
    pydantic \
    pydantic-settings \
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
COPY --chown=appuser:appuser run_server.py ./
COPY --chown=appuser:appuser pyproject.toml ./

# Environment variables
ENV PYTHONUNBUFFERED=1 \
    PYTHONDONTWRITEBYTECODE=1 \
    OUTPUT_DIR=/app/outputs \
    S3_ENDPOINT=https://dcs3.psn.co.id \
    S3_BUCKET_NAME= \
    S3_REGION= \
    S3_ACCESS_KEY= \
    S3_SECRET_KEY= \
    FILE_RETENTION_HOURS=24 \
    RATE_LIMIT_REQUESTS=20 \
    RATE_LIMIT_WINDOW=60 \
    MCP_TRANSPORT=stdio \
    LOCALE=en_US \
    DEBUG=false

# Health check
HEALTHCHECK --interval=30s --timeout=5s --start-period=10s --retries=3 \
    CMD python -c "import src.server; print('healthy')" || exit 1

# Default command (stdio mode)
CMD ["python", "-m", "src.server"]
