#!/usr/bin/env python
"""Wrapper script to run MCP Office Server with correct PYTHONPATH."""

import sys
from pathlib import Path

# Add project root to Python path
project_root = Path(__file__).parent.resolve()
if str(project_root) not in sys.path:
    sys.path.insert(0, str(project_root))


def main():
    """Entry point for console script."""
    import asyncio
    from src.server import main as server_main
    asyncio.run(server_main())


if __name__ == "__main__":
    main()