from __future__ import annotations

import logging
import re
from html.parser import HTMLParser

logger = logging.getLogger(__name__)


class _HTMLStripper(HTMLParser):
    """Minimal HTML-to-text converter."""

    def __init__(self) -> None:
        super().__init__()
        self._parts: list[str] = []

    def handle_data(self, data: str) -> None:
        self._parts.append(data)

    def handle_starttag(self, tag: str, attrs: list) -> None:
        # Insert line breaks for block elements so paragraphs don't merge
        if tag in ("p", "br", "div", "li", "tr", "h1", "h2", "h3", "h4", "h5", "h6"):
            self._parts.append("\n")

    def get_text(self) -> str:
        text = "".join(self._parts)
        # Collapse runs of blank lines to at most two
        text = re.sub(r"\n{3,}", "\n\n", text)
        return text.strip()


def strip_html(html: str) -> str:
    """Convert an HTML string to plain text."""
    if not html:
        return ""
    stripper = _HTMLStripper()
    try:
        stripper.feed(html)
        return stripper.get_text()
    except Exception as exc:
        # Log the failure so it doesn't silently mask problems
        logger.warning("HTML parsing failed, falling back to regex strip: %s", exc)
        # Fallback: crude tag removal
        return re.sub(r"<[^>]+>", "", html).strip()


def build_recipients(addresses: str) -> list[dict]:
    """Convert a comma-separated address string to a Graph API recipient list.

    Supports plain addresses ('a@b.com') and display-name format ('Name <a@b.com>').
    """
    result = []
    for raw in addresses.split(","):
        raw = raw.strip()
        if not raw:
            continue
        # Parse 'Display Name <email@domain.com>'
        match = re.match(r'^(.+?)\s*<(.+?)>$', raw)
        if match:
            name, addr = match.group(1).strip(), match.group(2).strip()
        else:
            name, addr = raw, raw
        result.append({"emailAddress": {"name": name, "address": addr}})
    return result


def truncate(text: str, max_len: int) -> str:
    """Truncate a string to max_len characters, appending '...' if cut."""
    if len(text) <= max_len:
        return text
    return text[:max_len] + "..."


# ---------------------------------------------------------------------------
# OData parameter sanitization
# ---------------------------------------------------------------------------

# Characters that have no place in safe OData identifiers / simple field lists.
# We don't attempt to parse the full OData grammar; instead we block characters
# that are only needed for injection payloads.
_ODATA_BLOCKED = frozenset(";\\`\x00")

# Reasonable upper bounds to prevent oversized query strings
_ODATA_MAX_LEN = 500


def sanitize_odata_param(value: str, name: str = "parameter") -> str:
    """Return *value* unchanged if it looks safe, otherwise raise ValueError.

    Rules:
    - Must not exceed _ODATA_MAX_LEN characters.
    - Must not contain characters in _ODATA_BLOCKED.
    """
    if not isinstance(value, str):
        raise ValueError(f"OData {name} must be a string")
    if len(value) > _ODATA_MAX_LEN:
        raise ValueError(
            f"OData {name} exceeds maximum allowed length of {_ODATA_MAX_LEN} characters"
        )
    for ch in value:
        if ch in _ODATA_BLOCKED:
            raise ValueError(
                f"OData {name} contains a disallowed character: {ch!r}"
            )
    return value
