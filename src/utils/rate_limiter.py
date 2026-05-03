"""Rate limiting with sliding window algorithm for per-user generation limits."""

from __future__ import annotations

import os
import time
from collections import defaultdict
from typing import Optional

from src.utils.logger import get_logger

log = get_logger("rate_limiter")


class RateLimiter:
    """Sliding window rate limiter for per-user request control.

    Features:
    - Per-user tracking with sliding window
    - Configurable limits via environment variables
    - Graceful rejection with retry-after information
    """

    def __init__(
        self,
        max_requests: Optional[int] = None,
        window_seconds: int = 60,
    ):
        """Initialize rate limiter.

        Args:
            max_requests: Max requests per window per user. Defaults to RATE_LIMIT_REQUESTS env var or 20.
            window_seconds: Time window in seconds. Defaults to RATE_LIMIT_WINDOW or 60.
        """
        self.max_requests = max_requests or int(os.environ.get("RATE_LIMIT_REQUESTS", "20"))
        self.window_seconds = window_seconds or int(os.environ.get("RATE_LIMIT_WINDOW", "60"))
        self._requests: dict[str, list[float]] = defaultdict(list)

    def is_allowed(self, user_id: str) -> tuple[bool, dict]:
        """Check if a request from the given user is allowed.

        Args:
            user_id: Unique user identifier (e.g., session ID).

        Returns:
            Tuple of (allowed, info_dict). Info dict contains:
            - remaining: Number of remaining requests in window
            - limit: Max requests per window
            - reset_seconds: Seconds until the window resets
            - retry_after: Seconds to wait before retrying (if not allowed)
        """
        now = time.time()
        window_start = now - self.window_seconds

        # Clean old entries for this user
        self._requests[user_id] = [
            ts for ts in self._requests[user_id] if ts > window_start
        ]

        current_count = len(self._requests[user_id])
        remaining = max(0, self.max_requests - current_count)

        # Calculate when the oldest request in window will expire
        if self._requests[user_id]:
            oldest = min(self._requests[user_id])
            reset_seconds = max(0, oldest + self.window_seconds - now)
        else:
            reset_seconds = 0

        info = {
            "remaining": remaining,
            "limit": self.max_requests,
            "reset_seconds": round(reset_seconds, 1),
        }

        if current_count >= self.max_requests:
            info["retry_after"] = round(reset_seconds, 1)
            log.warning(f"Rate limit exceeded for user {user_id}")
            return False, info

        # Record this request
        self._requests[user_id].append(now)
        info["remaining"] = remaining - 1  # Decrement after recording
        return True, info

    def get_user_stats(self, user_id: str) -> dict:
        """Get current rate limit stats for a user.

        Args:
            user_id: Unique user identifier.

        Returns:
            Dictionary with user rate limit statistics.
        """
        now = time.time()
        window_start = now - self.window_seconds

        current_requests = [
            ts for ts in self._requests.get(user_id, []) if ts > window_start
        ]

        return {
            "user_id": user_id,
            "current_requests": len(current_requests),
            "max_requests": self.max_requests,
            "remaining": max(0, self.max_requests - len(current_requests)),
            "window_seconds": self.window_seconds,
        }

    def reset_user(self, user_id: str) -> None:
        """Reset rate limit for a specific user.

        Args:
            user_id: Unique user identifier.
        """
        self._requests.pop(user_id, None)
        log.info(f"Rate limit reset for user {user_id}")

    def cleanup(self) -> None:
        """Remove stale entries for all users."""
        now = time.time()
        window_start = now - self.window_seconds

        stale_users = [
            user_id
            for user_id, timestamps in self._requests.items()
            if not any(ts > window_start for ts in timestamps)
        ]

        for user_id in stale_users:
            del self._requests[user_id]

        if stale_users:
            log.debug(f"Cleaned up {len(stale_users)} stale rate limit entries")