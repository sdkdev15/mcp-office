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
    starlette \
    boto3 \
    numpy \
    scipy \
    pandas \
    Pillow \
    requests \
    pypdf

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
    PYTHONDONTWRITEBYTECODE=1 

# Health check
HEALTHCHECK --interval=30s --timeout=5s --start-period=10s --retries=3 \
    CMD python -c "import src.server; print('healthy')" || exit 1

# Default command (stdio mode)
CMD ["python", "-m", "src.server"]
