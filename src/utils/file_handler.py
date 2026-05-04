"""File handling utilities for saving, serving, and managing generated documents."""

from __future__ import annotations

import asyncio
import mimetypes
import uuid
from pathlib import Path
from datetime import datetime
from typing import Optional

from src.utils.formatting import human_readable_size
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

    def _upload_to_s3(self, filepath: Path, key: str) -> None:
        import os
        import boto3
        from botocore.config import Config
        
        endpoint = os.environ.get("S3_ENDPOINT")
        bucket = os.environ.get("S3_BUCKET_NAME")
        
        if not endpoint or not bucket:
            log.warning("S3 variables not fully set. Skipping S3 upload.")
            return

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
        
        log.info(f"Uploading {filepath.name} to s3://{bucket}/{key}")
        s3.upload_file(str(filepath), bucket, key)

    async def save_file(
        self,
        data: bytes,
        filename: str,
        session_id: Optional[str] = None,
    ) -> dict:
        """Save generated file to S3 and keep local gzip backup.

        Args:
            data: File content as bytes.
            filename: Desired filename.
            session_id: Optional session ID for isolation.

        Returns:
            Dictionary with file metadata.
        """
        import os
        import gzip
        from datetime import datetime

        base_session = session_id or "default"
        # Add date partition if not already present
        if not (len(base_session) > 8 and base_session[:8].isdigit() and base_session.startswith("202")):
            date_str = datetime.now().strftime("%Y%m%d")
            effective_session = f"{date_str}{base_session}"
        else:
            effective_session = base_session

        session_dir = self.get_session_dir(effective_session)
        filepath = session_dir / filename

        # 1. Write uncompressed file temporarily
        await asyncio.to_thread(filepath.write_bytes, data)
        file_size = filepath.stat().st_size
        mime_type = self._get_mime_type(filepath.suffix)

        # 2. Upload to S3
        s3_key = f"{effective_session}/{filename}"
        try:
            await asyncio.to_thread(self._upload_to_s3, filepath, s3_key)
        except Exception as e:
            log.error(f"Failed to upload {filename} to S3: {e}")

        # 3. Compress to .gz and clean up uncompressed
        gz_filepath = filepath.with_name(f"{filename}.gz")
        
        def _compress():
            with filepath.open('rb') as f_in:
                with gzip.open(gz_filepath, 'wb') as f_out:
                    f_out.writelines(f_in)
            filepath.unlink() # delete uncompressed

        await asyncio.to_thread(_compress)
        log.info(f"File uploaded to S3 and compressed locally: {gz_filepath.name} ({gz_filepath.stat().st_size} bytes)")

        return {
            "filename": filename,
            "filepath": str(gz_filepath),
            "session_id": effective_session,
            "size_bytes": file_size,
            "mime_type": mime_type,
            "created_at": datetime.now().isoformat(),
            "resource_uri": f"/files/{effective_session}/{filename}",
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
        """Convert bytes to human-readable size string."""
        return human_readable_size(size_bytes)