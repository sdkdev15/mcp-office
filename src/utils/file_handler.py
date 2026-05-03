"""File handling utilities for saving, serving, and managing generated documents."""

from __future__ import annotations

import asyncio
import mimetypes
import uuid
from pathlib import Path
from datetime import datetime
from typing import Optional

from src.utils.logger import get_logger

log = get_logger("file_handler")


class FileHandler:
    """Handles file operations for generated documents with per-session isolation."""

    def __init__(self, output_dir: str = "outputs"):
        self.output_dir = Path(output_dir)
        self.output_dir.mkdir(parents=True, exist_ok=True)
        self._mime_types = {
            ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            ".ods": "application/vnd.oasis.opendocument.spreadsheet",
            ".csv": "text/csv",
            ".pptx": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
            ".odp": "application/vnd.oasis.opendocument.presentation",
            ".docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            ".odt": "application/vnd.oasis.opendocument.text",
            ".pdf": "application/pdf",
        }

    def get_session_dir(self, session_id: Optional[str] = None) -> Path:
        """Get or create a session-specific output directory.

        Args:
            session_id: Optional session ID. If not provided, generates a unique one.

        Returns:
            Path to the session directory.
        """
        if session_id is None:
            session_id = str(uuid.uuid4())[:8]
        session_dir = self.output_dir / session_id
        session_dir.mkdir(parents=True, exist_ok=True)
        return session_dir

    def generate_filename(self, base_name: str, extension: str) -> str:
        """Generate a unique filename with timestamp.

        Args:
            base_name: Base name for the file.
            extension: File extension (e.g., '.xlsx').

        Returns:
            Unique filename string.
        """
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        unique_id = uuid.uuid4().hex[:6]
        safe_name = base_name.replace(" ", "_").replace("/", "_")
        return f"{safe_name}_{timestamp}_{unique_id}{extension}"

    async def save_file(
        self,
        data: bytes,
        filename: str,
        session_id: Optional[str] = None,
    ) -> dict:
        """Save generated file to disk.

        Args:
            data: File content as bytes.
            filename: Desired filename.
            session_id: Optional session ID for isolation.

        Returns:
            Dictionary with file metadata.
        """
        session_dir = self.get_session_dir(session_id)
        filepath = session_dir / filename

        await asyncio.to_thread(filepath.write_bytes, data)

        file_size = filepath.stat().st_size
        mime_type = self._get_mime_type(filepath.suffix)

        log.info(f"File saved: {filepath.name} ({file_size} bytes)")

        return {
            "filename": filename,
            "filepath": str(filepath),
            "session_id": session_id or "default",
            "size_bytes": file_size,
            "mime_type": mime_type,
            "created_at": datetime.now().isoformat(),
            "resource_uri": f"file://{self.output_dir}/{session_id or 'default'}/{filename}",
        }

    async def read_file(self, filepath: str) -> bytes:
        """Read file content from disk.

        Args:
            filepath: Path to the file.

        Returns:
            File content as bytes.
        """
        path = Path(filepath)
        if not path.exists():
            raise FileNotFoundError(f"File not found: {filepath}")
        return await asyncio.to_thread(path.read_bytes)

    def _get_mime_type(self, extension: str) -> str:
        """Get MIME type for file extension.

        Args:
            extension: File extension (e.g., '.xlsx').

        Returns:
            MIME type string.
        """
        return self._mime_types.get(extension.lower(), "application/octet-stream")

    def list_session_files(self, session_id: str) -> list[dict]:
        """List all files in a session directory.

        Args:
            session_id: Session ID to list files for.

        Returns:
            List of file metadata dictionaries.
        """
        session_dir = self.output_dir / session_id
        if not session_dir.exists():
            return []

        files = []
        for filepath in session_dir.iterdir():
            if filepath.is_file():
                files.append({
                    "filename": filepath.name,
                    "size_bytes": filepath.stat().st_size,
                    "mime_type": self._get_mime_type(filepath.suffix),
                    "created_at": datetime.fromtimestamp(filepath.stat().st_ctime).isoformat(),
                    "resource_uri": f"file://{self.output_dir}/{session_id}/{filepath.name}",
                })

        return files

    def get_file_preview(self, filepath: str) -> dict:
        """Generate a text preview/summary for a file.

        Args:
            filepath: Path to the file.

        Returns:
            Preview information dictionary.
        """
        path = Path(filepath)
        if not path.exists():
            raise FileNotFoundError(f"File not found: {filepath}")

        stat = path.stat()
        extension = path.suffix.lower()

        preview = {
            "filename": path.name,
            "size_bytes": stat.st_size,
            "size_human": self._human_readable_size(stat.st_size),
            "extension": extension,
            "created_at": datetime.fromtimestamp(stat.st_ctime).isoformat(),
        }

        if extension == ".xlsx":
            preview["type"] = "spreadsheet"
        elif extension in (".ods", ".csv"):
            preview["type"] = "spreadsheet"
        elif extension == ".pptx":
            preview["type"] = "presentation"
        elif extension == ".odp":
            preview["type"] = "presentation"
        elif extension == ".docx":
            preview["type"] = "document"
        elif extension == ".odt":
            preview["type"] = "document"
        else:
            preview["type"] = "unknown"

        return preview

    @staticmethod
    def _human_readable_size(size_bytes: int) -> str:
        """Convert bytes to human-readable size string.

        Args:
            size_bytes: Size in bytes.

        Returns:
            Human-readable size string.
        """
        for unit in ["B", "KB", "MB", "GB"]:
            if size_bytes < 1024:
                return f"{size_bytes:.1f} {unit}"
            size_bytes /= 1024
        return f"{size_bytes:.1f} TB"


# Global file handler instance
file_handler = FileHandler()