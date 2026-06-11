from __future__ import annotations

import logging
import signal
import threading
import time
from typing import Any

import httpx

from outlook_mcp.config import GRAPH_BASE_URL, OUTLOOK_PROXY, OUTLOOK_TOKEN, VERIFY_SSL

logger = logging.getLogger(__name__)

# How long to wait for a Retry-After header before using the default back-off
_MAX_RETRY_AFTER_SECS = 60


class GraphError(Exception):
    """Raised when the Microsoft Graph API returns an error."""


class TokenExpiredError(GraphError):
    """Raised specifically when the token is expired or invalid (HTTP 401)."""


def _headers() -> dict[str, str]:
    return {
        "Authorization": f"Bearer {OUTLOOK_TOKEN}",
        "Accept": "application/json",
    }


def _parse_error(response: httpx.Response) -> str:
    """Build a human-readable error string from a Graph API error response.

    NOTE: Only the Graph-supplied error *code* and *message* fields are forwarded
    to callers, never the raw response body, to avoid leaking sensitive debug data.
    """
    codes: dict[int, str] = {
        400: "Bad request — check your input parameters",
        401: "Unauthorized — your OUTLOOK_TOKEN is missing or expired",
        403: "Forbidden — your token lacks the required Graph API permissions",
        404: "Not found — the email/folder/resource does not exist",
        429: "Rate limited — too many requests, try again shortly",
    }
    base = codes.get(response.status_code, f"Graph API error {response.status_code}")
    try:
        body = response.json()
        err = body.get("error", {})
        # Only surface the structured error fields, not the raw body
        if err.get("code"):
            base = f"[{err['code']}] {base}"
        if err.get("message"):
            base += f": {err['message']}"
    except Exception:
        # No structured body — don't append raw text (could leak sensitive data)
        pass
    return base


class GraphClient:
    """httpx client wrapping Microsoft Graph API with retry and 401-detection logic."""

    def __init__(self) -> None:
        if not OUTLOOK_TOKEN:
            raise GraphError("OUTLOOK_TOKEN environment variable is not set")

        transport_mounts: dict | None = None
        if OUTLOOK_PROXY:
            transport_mounts = {
                "http://": httpx.HTTPTransport(proxy=OUTLOOK_PROXY),
                "https://": httpx.HTTPTransport(proxy=OUTLOOK_PROXY),
            }

        self._client = httpx.Client(
            base_url=GRAPH_BASE_URL,
            headers=_headers(),
            verify=VERIFY_SSL,
            timeout=httpx.Timeout(60.0, connect=10.0),
            limits=httpx.Limits(max_connections=20, max_keepalive_connections=10),
            mounts=transport_mounts,
            trust_env=False,
        )
        logger.info("GraphClient initialised (SSL verify=%s)", VERIFY_SSL)

    def _request(self, method: str, path: str, **kwargs: Any) -> Any:
        max_retries = 3
        for attempt in range(max_retries):
            logger.debug("Graph request attempt %d: %s %s", attempt + 1, method, path)
            try:
                resp = self._client.request(method, path, **kwargs)
            except httpx.TransportError as exc:
                logger.warning("Transport error on attempt %d: %s", attempt + 1, exc)
                if attempt < max_retries - 1:
                    time.sleep(attempt + 1)
                    continue
                raise GraphError(f"Connection error: {exc}") from exc

            logger.debug("Graph response: %d %s %s", resp.status_code, method, path)

            # 401 — token is expired or invalid; raise a dedicated error immediately
            if resp.status_code == 401:
                logger.error("Graph API returned 401 — token is expired or invalid")
                raise TokenExpiredError(
                    "OUTLOOK_TOKEN is expired or invalid. "
                    "Please refresh your token and restart the server."
                )

            if resp.status_code == 429 or resp.status_code >= 500:
                if attempt < max_retries - 1:
                    # Respect Retry-After header when present
                    retry_after_raw = resp.headers.get("Retry-After")
                    try:
                        retry_after = min(float(retry_after_raw), _MAX_RETRY_AFTER_SECS)
                    except (TypeError, ValueError):
                        retry_after = (attempt + 1) * (2 if resp.status_code == 429 else 1)
                    logger.warning(
                        "Retryable response %d on attempt %d; sleeping %.1fs",
                        resp.status_code, attempt + 1, retry_after,
                    )
                    time.sleep(retry_after)
                    continue
                raise GraphError(_parse_error(resp))

            if resp.status_code >= 400:
                raise GraphError(_parse_error(resp))
            if resp.status_code == 204:
                return {}
            return resp.json()

        raise GraphError(f"Max retries exceeded for {method} {path}")

    def get(self, path: str) -> Any:
        return self._request("GET", path)

    def post(self, path: str, payload: Any = None) -> Any:
        return self._request("POST", path, json=payload)

    def patch(self, path: str, payload: Any = None) -> Any:
        return self._request("PATCH", path, json=payload)

    def close(self) -> None:
        logger.info("Closing GraphClient HTTP connection pool")
        self._client.close()


# ---------------------------------------------------------------------------
# Module-level singleton — created lazily on first use
# ---------------------------------------------------------------------------
_client: GraphClient | None = None
_client_lock = threading.Lock()


def get_client() -> GraphClient:
    global _client
    if _client is None:
        with _client_lock:
            if _client is None:
                _client = GraphClient()  # raises GraphError if OUTLOOK_TOKEN is missing
    return _client


def _shutdown_client() -> None:
    """Close the singleton client on process exit / signal."""
    global _client
    with _client_lock:
        if _client is not None:
            _client.close()
            _client = None


def _handle_signal(signum: int, frame: Any) -> None:  # noqa: ARG001
    logger.info("Received signal %d — shutting down", signum)
    _shutdown_client()
    raise SystemExit(0)


# Register graceful shutdown handlers so Docker SIGTERM / Ctrl-C close the pool
signal.signal(signal.SIGTERM, _handle_signal)
signal.signal(signal.SIGINT, _handle_signal)
