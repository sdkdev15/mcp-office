"""File cleanup system with startup cleanup and background thread for production use."""

from __future__ import annotations

import asyncio
import os
import threading
import time
from datetime import datetime, timedelta
from pathlib import Path
from typing import Optional

from src.utils.formatting import human_readable_size
from src.utils.logger import get_logger

log = get_logger("cleanup")


class FileCleanup:
    """Manages automatic cleanup of old generated files.

    Features:
    - Startup cleanup: Delete files older than retention period on server start
    - Background cleanup: Runs every hour to delete old files
    - Configurable retention via FILE_RETENTION_HOURS environment variable
    - Per-session directory isolation
    """

    def __init__(
        self,
        output_dir: str = "outputs",
        retention_hours: Optional[int] = None,
        cleanup_interval_hours: int = 1,
    ):
        self.output_dir = Path(output_dir)
        self.retention_hours = retention_hours or int(os.environ.get("FILE_RETENTION_HOURS", "24"))
        self.cleanup_interval = cleanup_interval_hours * 3600  # Convert to seconds
        self._running = False
        self._cleanup_thread: Optional[threading.Thread] = None

    def start(self) -> None:
        """Start the background cleanup thread and perform initial cleanup."""
        log.info(f"Starting file cleanup system (retention: {self.retention_hours}h, interval: {self.cleanup_interval // 3600}h)")

        # Perform initial cleanup on startup
        self.cleanup_old_files()

        # Start background thread
        self._running = True
        self._cleanup_thread = threading.Thread(target=self._background_cleanup, daemon=True)
        self._cleanup_thread.start()
        log.info("Background cleanup thread started")

    def stop(self) -> None:
        """Stop the background cleanup thread."""
        self._running = False
        if self._cleanup_thread:
            self._cleanup_thread.join(timeout=5)
        log.info("Cleanup system stopped")

    def cleanup_old_files(self) -> int:
        """Delete files older than the retention period.

        Returns:
            Number of files deleted.
        """
        if not self.output_dir.exists():
            return 0

        cutoff_time = datetime.now() - timedelta(hours=self.retention_hours)
        deleted_count = 0

        for session_dir in self.output_dir.iterdir():
            if not session_dir.is_dir():
                continue

            for filepath in session_dir.iterdir():
                if not filepath.is_file():
                    continue

                file_mtime = datetime.fromtimestamp(filepath.stat().st_mtime)
                if file_mtime < cutoff_time:
                    try:
                        filepath.unlink()
                        deleted_count += 1
                        log.debug(f"Deleted old file: {filepath.name}")
                    except OSError as e:
                        log.error(f"Failed to delete {filepath.name}: {e}")

            # Remove empty session directories
            try:
                if not any(session_dir.iterdir()):
                    session_dir.rmdir()
                    log.debug(f"Removed empty session directory: {session_dir.name}")
            except OSError:
                pass

        if deleted_count > 0:
            log.info(f"Cleanup complete: deleted {deleted_count} files older than {self.retention_hours}h")

        return deleted_count

    def get_storage_stats(self) -> dict:
        """Get current storage usage statistics.

        Returns:
            Dictionary with storage statistics.
        """
        total_files = 0
        total_size = 0
        sessions = 0

        if self.output_dir.exists():
            for session_dir in self.output_dir.iterdir():
                if not session_dir.is_dir():
                    continue
                sessions += 1

                for filepath in session_dir.iterdir():
                    if filepath.is_file():
                        total_files += 1
                        total_size += filepath.stat().st_size

        return {
            "total_files": total_files,
            "total_size_bytes": total_size,
            "total_size_human": self._human_readable_size(total_size),
            "active_sessions": sessions,
            "retention_hours": self.retention_hours,
        }

    def _background_cleanup(self) -> None:
        """Background thread that periodically cleans up old files."""
        while self._running:
            time.sleep(self.cleanup_interval)
            if self._running:
                try:
                    self.cleanup_old_files()
                except Exception as e:
                    log.error(f"Background cleanup error: {e}")

    @staticmethod
    def _human_readable_size(size_bytes: int) -> str:
        """Convert bytes to human-readable size string."""
        return human_readable_size(size_bytes)