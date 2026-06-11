"""Tests for outlook_mcp.client — GraphClient, retry, 401 handling, shutdown."""

from __future__ import annotations

import os
import time
from unittest.mock import patch

import httpx
import pytest
import respx

os.environ.setdefault("OUTLOOK_TOKEN", "test-token")
os.environ.setdefault("OUTLOOK_VERIFY_SSL", "false")

import outlook_mcp.client as client_mod  # noqa: E402

from outlook_mcp.client import (  # noqa: E402
    GraphClient,
    GraphError,
    TokenExpiredError,
    _shutdown_client,
    get_client,
)


# ---------------------------------------------------------------------------
# GraphClient initialisation
# ---------------------------------------------------------------------------

class TestGraphClientInit:
    def test_raises_if_token_missing(self):
        """GraphClient raises immediately when OUTLOOK_TOKEN is empty."""
        with patch.object(client_mod, "OUTLOOK_TOKEN", ""):
            with pytest.raises(GraphError, match="OUTLOOK_TOKEN"):
                GraphClient()

    def test_raises_for_missing_token_at_get_client(self):
        """get_client() propagates GraphError when token is absent."""
        with patch.object(client_mod, "OUTLOOK_TOKEN", ""):
            with pytest.raises(GraphError, match="OUTLOOK_TOKEN"):
                get_client()


# ---------------------------------------------------------------------------
# Successful requests
# ---------------------------------------------------------------------------

class TestGraphClientRequests:
    @respx.mock
    def test_get_returns_json(self):
        respx.get("https://graph.microsoft.com/v1.0/me/messages").mock(
            return_value=httpx.Response(200, json={"value": [{"id": "abc"}]})
        )
        client = GraphClient()
        result = client.get("/me/messages")
        assert result == {"value": [{"id": "abc"}]}
        client.close()

    @respx.mock
    def test_post_returns_json(self):
        respx.post("https://graph.microsoft.com/v1.0/me/sendMail").mock(
            return_value=httpx.Response(204)
        )
        client = GraphClient()
        result = client.post("/me/sendMail", {"message": {}})
        assert result == {}
        client.close()

    @respx.mock
    def test_204_returns_empty_dict(self):
        respx.patch("https://graph.microsoft.com/v1.0/me/messages/123").mock(
            return_value=httpx.Response(204)
        )
        client = GraphClient()
        result = client.patch("/me/messages/123", {"isRead": True})
        assert result == {}
        client.close()


# ---------------------------------------------------------------------------
# Error handling
# ---------------------------------------------------------------------------

class TestGraphClientErrors:
    @respx.mock
    def test_404_raises_graph_error(self):
        respx.get("https://graph.microsoft.com/v1.0/me/messages/missing").mock(
            return_value=httpx.Response(
                404, json={"error": {"code": "ErrorItemNotFound", "message": "Item not found"}}
            )
        )
        client = GraphClient()
        exc = None
        try:
            client.get("/me/messages/missing")
        except GraphError as e:
            exc = e
        finally:
            client.close()
        assert exc is not None, "Expected GraphError to be raised"
        assert "Not found" in str(exc)

    @respx.mock
    def test_401_raises_token_expired_error(self):
        respx.get("https://graph.microsoft.com/v1.0/me/messages").mock(
            return_value=httpx.Response(
                401, json={"error": {"code": "InvalidAuthenticationToken", "message": "Access token expired."}}
            )
        )
        client = GraphClient()
        exc = None
        try:
            client.get("/me/messages")
        except TokenExpiredError as e:
            exc = e
        finally:
            client.close()
        assert exc is not None, "Expected TokenExpiredError to be raised"
        assert "expired or invalid" in str(exc)

    @respx.mock
    def test_403_raises_graph_error(self):
        respx.get("https://graph.microsoft.com/v1.0/me/messages").mock(
            return_value=httpx.Response(403, json={"error": {"code": "AccessDenied", "message": "Forbidden"}})
        )
        client = GraphClient()
        exc = None
        try:
            client.get("/me/messages")
        except GraphError as e:
            exc = e
        finally:
            client.close()
        assert exc is not None, "Expected GraphError to be raised"
        assert "Forbidden" in str(exc)

    @respx.mock
    def test_transport_error_raises_graph_error(self):
        respx.get("https://graph.microsoft.com/v1.0/me/messages").mock(
            side_effect=httpx.ConnectError("connection refused")
        )
        client = GraphClient()
        exc = None
        with patch("time.sleep"):
            try:
                client.get("/me/messages")
            except GraphError as e:
                exc = e
            finally:
                client.close()
        assert exc is not None, "Expected GraphError to be raised"
        assert "Connection error" in str(exc)


# ---------------------------------------------------------------------------
# Retry / back-off
# ---------------------------------------------------------------------------

class TestGraphClientRetry:
    @respx.mock
    def test_retries_on_500_then_succeeds(self):
        route = respx.get("https://graph.microsoft.com/v1.0/me/messages")
        route.side_effect = [
            httpx.Response(500, json={}),
            httpx.Response(500, json={}),
            httpx.Response(200, json={"value": []}),
        ]
        client = GraphClient()
        with patch("time.sleep"):
            result = client.get("/me/messages")
        assert result == {"value": []}
        client.close()

    @respx.mock
    def test_respects_retry_after_header(self):
        """Verifies Retry-After header value is used for sleep duration."""
        slept: list[float] = []

        route = respx.get("https://graph.microsoft.com/v1.0/me/messages")
        route.side_effect = [
            httpx.Response(429, headers={"Retry-After": "5"}, json={}),
            httpx.Response(200, json={"value": []}),
        ]
        client = GraphClient()
        with patch("time.sleep", side_effect=lambda s: slept.append(s)):
            client.get("/me/messages")
        client.close()
        assert slept == [5.0], f"Expected sleep of 5s, got {slept}"

    @respx.mock
    def test_max_retries_exhausted_raises(self):
        respx.get("https://graph.microsoft.com/v1.0/me/messages").mock(
            return_value=httpx.Response(500, json={})
        )
        client = GraphClient()
        with patch("time.sleep"):
            exc = None
            try:
                client.get("/me/messages")
            except GraphError as e:
                exc = e
            finally:
                client.close()
        assert exc is not None, "Expected GraphError after exhausted retries"


# ---------------------------------------------------------------------------
# Singleton / shutdown
# ---------------------------------------------------------------------------

class TestSingleton:
    @respx.mock
    def test_get_client_returns_same_instance(self):
        respx.get("https://graph.microsoft.com/v1.0/me/messages").mock(
            return_value=httpx.Response(200, json={"value": []})
        )
        c1 = get_client()
        c2 = get_client()
        assert c1 is c2

    def test_shutdown_client_closes_and_nils_singleton(self):
        with respx.mock:
            respx.get("https://graph.microsoft.com/v1.0/me/messages").mock(
                return_value=httpx.Response(200, json={"value": []})
            )
            get_client()

        assert client_mod._client is not None
        _shutdown_client()
        assert client_mod._client is None

