"""Security utilities: audit trail, PII redaction, and input sanitization."""

from __future__ import annotations

import re
import hashlib
import json
import os
from datetime import datetime
from pathlib import Path
from typing import Any, Optional

from src.utils.logger import get_logger

log = get_logger("security")

# PII patterns
PII_PATTERNS = [
    # Email
    (r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}', '[EMAIL]'),
    # Phone numbers (international)
    (r'\b[\+]?[1-9][0-9]{7,14}\b', '[PHONE]'),
    # Indonesian phone
    (r'\b08[1-9][0-9]{8,11}\b', '[PHONE]'),
    # Credit card (basic)
    (r'\b(?:\d{4}[-\s]?){3}\d{4}\b', '[CARD]'),
    # KTP (Indonesian ID - 16 digits)
    (r'\b\d{16}\b', '[KTP]'),
    # NPWP (Indonesian tax ID)
    (r'\b\d{2}[\.\-]\d{3}[\.\-]\d{3}[\.\-]\d{4}[\.\-]\d{2}[\.\-]\d{3}\b', '[NPWP]'),
]


class AuditTrail:
    """Simple file-based audit trail for compliance."""

    def __init__(self, audit_dir: str = "outputs/.audit"):
        self.audit_dir = Path(audit_dir)
        self.audit_dir.mkdir(parents=True, exist_ok=True)

    def log_action(
        self,
        action: str,
        user_id: str,
        details: Optional[dict] = None,
    ) -> str:
        """Log an action to the audit trail.

        Args:
            action: Action type (e.g., 'file_generated', 'file_deleted').
            user_id: User/session identifier.
            details: Optional action details.

        Returns:
            Audit entry ID.
        """
        timestamp = datetime.now().isoformat()
        entry_id = hashlib.sha256(f"{timestamp}{user_id}{action}".encode()).hexdigest()[:12]

        entry = {
            "id": entry_id,
            "timestamp": timestamp,
            "action": action,
            "user_id": user_id,
            "details": details or {},
            "ip_address": os.environ.get("CLIENT_IP", "unknown"),
        }

        # Write to daily audit file
        date_str = datetime.now().strftime("%Y-%m-%d")
        audit_file = self.audit_dir / f"audit_{date_str}.jsonl"

        with open(audit_file, "a", encoding="utf-8") as f:
            f.write(json.dumps(entry) + "\n")

        log.info(f"Audit: {action} by {user_id} ({entry_id})")
        return entry_id

    def get_entries(
        self,
        user_id: Optional[str] = None,
        action: Optional[str] = None,
        limit: int = 100,
    ) -> list[dict]:
        """Query audit entries.

        Args:
            user_id: Filter by user ID.
            action: Filter by action type.
            limit: Max entries to return.

        Returns:
            List of audit entries.
        """
        entries = []
        for audit_file in sorted(self.audit_dir.glob("audit_*.jsonl"), reverse=True):
            with open(audit_file, "r", encoding="utf-8") as f:
                for line in f:
                    if not line.strip():
                        continue
                    entry = json.loads(line)
                    if user_id and entry["user_id"] != user_id:
                        continue
                    if action and entry["action"] != action:
                        continue
                    entries.append(entry)
                    if len(entries) >= limit:
                        return entries
        return entries


class PIIRedactor:
    """Redact Personally Identifiable Information from text."""

    def __init__(self):
        self._compiled_patterns = [
            (re.compile(pattern), replacement)
            for pattern, replacement in PII_PATTERNS
        ]

    def redact(self, text: str) -> str:
        """Redact PII from text.

        Args:
            text: Input text.

        Returns:
            Text with PII replaced by placeholders.
        """
        if not text:
            return text

        redacted = text
        for pattern, replacement in self._compiled_patterns:
            redacted = pattern.sub(replacement, redacted)

        return redacted

    def redact_data(self, data: Any) -> Any:
        """Recursively redact PII from data structures.

        Args:
            data: Data to redact (str, list, dict).

        Returns:
            Redacted data.
        """
        if isinstance(data, str):
            return self.redact(data)
        elif isinstance(data, list):
            return [self.redact_data(item) for item in data]
        elif isinstance(data, dict):
            return {k: self.redact_data(v) for k, v in data.items()}
        return data


class InputSanitizer:
    """Sanitize user inputs to prevent injection attacks."""

    # Dangerous patterns
    DANGEROUS_PATTERNS = [
        r'<script[^>]*>.*?</script>',  # Script tags
        r'javascript:',  # JavaScript protocol
        r'on\w+\s*=',  # Event handlers
        r'\.\./',  # Path traversal
        r';\s*(?:rm|del|drop|truncate)',  # Command injection
    ]

    def __init__(self):
        self._compiled = [
            re.compile(p, re.IGNORECASE) for p in self.DANGEROUS_PATTERNS
        ]

    def sanitize(self, text: str) -> str:
        """Remove dangerous patterns from text.

        Args:
            text: Input text.

        Returns:
            Sanitized text.
        """
        if not text:
            return text

        sanitized = text
        for pattern in self._compiled:
            sanitized = pattern.sub('', sanitized)

        return sanitized

    def is_safe(self, text: str) -> bool:
        """Check if text is safe (no dangerous patterns).

        Args:
            text: Input text.

        Returns:
            True if safe, False otherwise.
        """
        return not any(p.search(text) for p in self._compiled)


# Global instances
audit_trail = AuditTrail()
pii_redactor = PIIRedactor()
input_sanitizer = InputSanitizer()