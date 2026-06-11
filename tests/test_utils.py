"""Tests for outlook_mcp.utils — HTML stripping, recipient building, OData sanitization."""

from __future__ import annotations

import pytest

from outlook_mcp.utils import build_recipients, sanitize_odata_param, strip_html, truncate


# ---------------------------------------------------------------------------
# strip_html
# ---------------------------------------------------------------------------

class TestStripHtml:
    def test_basic_html(self):
        assert strip_html("<p>Hello <b>world</b></p>") == "Hello world"

    def test_empty_string(self):
        assert strip_html("") == ""

    def test_plain_text_unchanged(self):
        assert strip_html("no tags here") == "no tags here"

    def test_br_becomes_newline(self):
        result = strip_html("line1<br>line2")
        assert "line1" in result
        assert "line2" in result

    def test_collapses_excess_blank_lines(self):
        result = strip_html("<p>a</p><p></p><p></p><p></p><p>b</p>")
        assert "\n\n\n" not in result

    def test_none_like_empty_returns_empty(self):
        assert strip_html("") == ""


# ---------------------------------------------------------------------------
# build_recipients
# ---------------------------------------------------------------------------

class TestBuildRecipients:
    def test_single_plain_address(self):
        result = build_recipients("alice@example.com")
        assert result == [{"emailAddress": {"name": "alice@example.com", "address": "alice@example.com"}}]

    def test_display_name_format(self):
        result = build_recipients("Alice Smith <alice@example.com>")
        assert result[0]["emailAddress"]["name"] == "Alice Smith"
        assert result[0]["emailAddress"]["address"] == "alice@example.com"

    def test_multiple_addresses(self):
        result = build_recipients("a@b.com, c@d.com")
        assert len(result) == 2

    def test_empty_string_returns_empty_list(self):
        assert build_recipients("") == []

    def test_whitespace_only_entry_skipped(self):
        result = build_recipients("a@b.com,   , c@d.com")
        assert len(result) == 2

    def test_mixed_formats(self):
        result = build_recipients("Alice <alice@example.com>, bob@example.com")
        assert result[0]["emailAddress"]["name"] == "Alice"
        assert result[1]["emailAddress"]["name"] == "bob@example.com"


# ---------------------------------------------------------------------------
# truncate
# ---------------------------------------------------------------------------

class TestTruncate:
    def test_short_string_unchanged(self):
        assert truncate("hi", 10) == "hi"

    def test_exact_length_unchanged(self):
        assert truncate("hello", 5) == "hello"

    def test_long_string_truncated(self):
        result = truncate("hello world", 5)
        assert result == "hello..."

    def test_empty_string(self):
        assert truncate("", 10) == ""


# ---------------------------------------------------------------------------
# sanitize_odata_param
# ---------------------------------------------------------------------------

class TestSanitizeOdataParam:
    def test_valid_filter_passes(self):
        val = "isRead eq false"
        assert sanitize_odata_param(val, "filter") == val

    def test_valid_orderby_passes(self):
        val = "receivedDateTime desc"
        assert sanitize_odata_param(val, "orderby") == val

    def test_semicolon_blocked(self):
        with pytest.raises(ValueError, match="disallowed character"):
            sanitize_odata_param("isRead eq false; DROP TABLE", "filter")

    def test_backslash_blocked(self):
        with pytest.raises(ValueError, match="disallowed character"):
            sanitize_odata_param("subject eq 'test\\x00'", "filter")

    def test_backtick_blocked(self):
        with pytest.raises(ValueError, match="disallowed character"):
            sanitize_odata_param("`cmd`", "search")

    def test_null_byte_blocked(self):
        with pytest.raises(ValueError, match="disallowed character"):
            sanitize_odata_param("value\x00evil", "filter")

    def test_too_long_blocked(self):
        long_val = "a" * 501
        with pytest.raises(ValueError, match="maximum allowed length"):
            sanitize_odata_param(long_val, "filter")

    def test_exactly_max_length_passes(self):
        val = "a" * 500
        assert sanitize_odata_param(val) == val

    def test_non_string_raises(self):
        with pytest.raises(ValueError, match="must be a string"):
            sanitize_odata_param(123)  # type: ignore[arg-type]
