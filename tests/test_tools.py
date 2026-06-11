"""Tests for email read/write tools using a mocked Graph API."""

from __future__ import annotations

import os

import httpx
import pytest
import respx

os.environ.setdefault("OUTLOOK_TOKEN", "test-token")
os.environ.setdefault("OUTLOOK_VERIFY_SSL", "false")

# Tools register themselves on import via @mcp.tool() — import after env is set
from outlook_mcp.tools.email_write import move_email, reply_email, send_email  # noqa: E402
from outlook_mcp.tools.emails import list_emails, read_email, search_emails  # noqa: E402


# ---------------------------------------------------------------------------
# list_emails
# ---------------------------------------------------------------------------

class TestListEmails:
    @respx.mock
    def test_returns_formatted_email_list(self):
        respx.get(url__regex=r".*/me/mailFolders/Inbox/messages.*").mock(
            return_value=httpx.Response(200, json={
                "value": [
                    {
                        "id": "msg1",
                        "subject": "Test Subject",
                        "from": {"emailAddress": {"name": "Alice", "address": "alice@example.com"}},
                        "receivedDateTime": "2026-06-10T10:00:00Z",
                        "isRead": False,
                        "hasAttachments": False,
                        "bodyPreview": "Preview text",
                    }
                ]
            })
        )
        result = list_emails(folder="Inbox", top=5)
        assert "Test Subject" in result
        assert "alice@example.com" in result
        assert "[UNREAD]" in result

    @respx.mock
    def test_top_capped_at_50(self):
        """Internally top > 50 is reset to 10; route must match any query."""
        route = respx.get(url__regex=r".*/me/mailFolders/Inbox/messages.*").mock(
            return_value=httpx.Response(200, json={"value": []})
        )
        list_emails(folder="Inbox", top=9999)
        assert route.called

    @respx.mock
    def test_graph_error_returns_error_string(self):
        respx.get(url__regex=r".*/me/mailFolders/Inbox/messages.*").mock(
            return_value=httpx.Response(403, json={"error": {"code": "AccessDenied", "message": "No access"}})
        )
        result = list_emails()
        assert result.startswith("ERROR:")

    def test_invalid_filter_returns_error(self):
        result = list_emails(filter="subject eq 'x'; DROP TABLE")
        assert result.startswith("ERROR:")

    def test_invalid_orderby_returns_error(self):
        result = list_emails(orderby="receivedDateTime`desc")
        assert result.startswith("ERROR:")


# ---------------------------------------------------------------------------
# read_email
# ---------------------------------------------------------------------------

class TestReadEmail:
    def test_missing_id_returns_error(self):
        result = read_email(id="")
        assert result.startswith("ERROR:")

    @respx.mock
    def test_returns_email_content(self):
        respx.get(url__regex=r".*/me/messages/msg1.*").mock(
            return_value=httpx.Response(200, json={
                "id": "msg1",
                "subject": "Hello World",
                "from": {"emailAddress": {"name": "Bob", "address": "bob@example.com"}},
                "toRecipients": [{"emailAddress": {"name": "Me", "address": "me@example.com"}}],
                "ccRecipients": [],
                "receivedDateTime": "2026-06-10T10:00:00Z",
                "importance": "normal",
                "hasAttachments": False,
                "body": {"contentType": "text", "content": "Body text here."},
            })
        )
        result = read_email(id="msg1")
        assert "Hello World" in result
        assert "bob@example.com" in result
        assert "Body text here." in result

    @respx.mock
    def test_html_body_stripped(self):
        respx.get(url__regex=r".*/me/messages/html1.*").mock(
            return_value=httpx.Response(200, json={
                "id": "html1",
                "subject": "HTML Email",
                "from": {"emailAddress": {"name": "X", "address": "x@x.com"}},
                "toRecipients": [],
                "ccRecipients": [],
                "receivedDateTime": "2026-06-10T10:00:00Z",
                "importance": "normal",
                "hasAttachments": False,
                "body": {"contentType": "html", "content": "<p>Hello <b>world</b></p>"},
            })
        )
        result = read_email(id="html1")
        assert "<p>" not in result
        assert "Hello" in result

    @respx.mock
    def test_token_expired_surfaced(self):
        respx.get(url__regex=r".*/me/messages/expired.*").mock(
            return_value=httpx.Response(401, json={"error": {"code": "InvalidAuthenticationToken"}})
        )
        result = read_email(id="expired")
        assert "ERROR:" in result
        assert "expired" in result.lower()


# ---------------------------------------------------------------------------
# search_emails
# ---------------------------------------------------------------------------

class TestSearchEmails:
    def test_empty_query_returns_error(self):
        result = search_emails(query="")
        assert result.startswith("ERROR:")

    def test_injection_query_returns_error(self):
        result = search_emails(query="test; DROP TABLE")
        assert result.startswith("ERROR:")

    @respx.mock
    def test_returns_results(self):
        respx.get(url__regex=r".*/me/messages.*").mock(
            return_value=httpx.Response(200, json={
                "value": [
                    {
                        "id": "s1",
                        "subject": "Budget Report",
                        "from": {"emailAddress": {"name": "Finance", "address": "finance@corp.com"}},
                        "receivedDateTime": "2026-06-01T09:00:00Z",
                        "bodyPreview": "Q2 numbers",
                        "isRead": True,
                        "hasAttachments": False,
                        "importance": "normal",
                    }
                ]
            })
        )
        result = search_emails(query="budget")
        assert "Budget Report" in result
        assert "finance@corp.com" in result


# ---------------------------------------------------------------------------
# send_email
# ---------------------------------------------------------------------------

class TestSendEmail:
    def test_missing_to_returns_error(self):
        result = send_email(to="", subject="Hello")
        assert result.startswith("ERROR:")

    def test_missing_subject_returns_error(self):
        result = send_email(to="a@b.com", subject="")
        assert result.startswith("ERROR:")

    @respx.mock
    def test_successful_send(self):
        respx.post("https://graph.microsoft.com/v1.0/me/sendMail").mock(
            return_value=httpx.Response(204)
        )
        result = send_email(to="alice@example.com", subject="Hi", body="Hello")
        assert "alice@example.com" in result
        assert "ERROR" not in result

    @respx.mock
    def test_graph_error_returns_error_string(self):
        respx.post("https://graph.microsoft.com/v1.0/me/sendMail").mock(
            return_value=httpx.Response(403, json={"error": {"code": "AccessDenied", "message": "No permission"}})
        )
        result = send_email(to="alice@example.com", subject="Test")
        assert result.startswith("ERROR:")


# ---------------------------------------------------------------------------
# reply_email
# ---------------------------------------------------------------------------

class TestReplyEmail:
    def test_missing_id_returns_error(self):
        result = reply_email(id="", comment="Thanks")
        assert result.startswith("ERROR:")

    def test_missing_comment_returns_error(self):
        result = reply_email(id="msg1", comment="")
        assert result.startswith("ERROR:")

    @respx.mock
    def test_reply_succeeds(self):
        respx.post(url__regex=r".*/me/messages/msg1/reply.*").mock(
            return_value=httpx.Response(204)
        )
        result = reply_email(id="msg1", comment="Got it!")
        assert "Reply" in result
        assert "ERROR" not in result

    @respx.mock
    def test_reply_all_succeeds(self):
        respx.post(url__regex=r".*/me/messages/msg2/replyAll.*").mock(
            return_value=httpx.Response(204)
        )
        result = reply_email(id="msg2", comment="Noted.", reply_all=True)
        assert "Reply-all" in result


# ---------------------------------------------------------------------------
# move_email
# ---------------------------------------------------------------------------

class TestMoveEmail:
    def test_missing_id_returns_error(self):
        result = move_email(id="", folder_id="Inbox")
        assert result.startswith("ERROR:")

    def test_missing_folder_returns_error(self):
        result = move_email(id="msg1", folder_id="")
        assert result.startswith("ERROR:")

    @respx.mock
    def test_move_succeeds(self):
        respx.post(url__regex=r".*/me/messages/msg1/move.*").mock(
            return_value=httpx.Response(200, json={"id": "msg1", "parentFolderId": "Archive"})
        )
        result = move_email(id="msg1", folder_id="Archive")
        assert "msg1" in result
        assert "Archive" in result
        assert "ERROR" not in result
