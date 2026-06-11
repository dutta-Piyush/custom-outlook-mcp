"""Shared pytest fixtures for the outlook-mcp test suite."""

from __future__ import annotations

import os

import pytest

# ---------------------------------------------------------------------------
# Ensure OUTLOOK_TOKEN is always set so config / client imports don't fail
# ---------------------------------------------------------------------------
os.environ.setdefault("OUTLOOK_TOKEN", "test-token-fixture")
os.environ.setdefault("OUTLOOK_VERIFY_SSL", "false")


@pytest.fixture(autouse=True)
def reset_graph_client():
    """Reset the module-level GraphClient singleton before and after every test.

    We deliberately do NOT use importlib.reload here because reloading the
    client module creates new class objects, which invalidates any class
    references already imported into test files (causing isinstance / except
    checks to silently fail).
    """
    import outlook_mcp.client as client_mod

    # Close any live client left over from a previous test
    with client_mod._client_lock:
        if client_mod._client is not None:
            try:
                client_mod._client.close()
            except Exception:
                pass
            client_mod._client = None
    yield
    # Tear-down: close again after the test
    with client_mod._client_lock:
        if client_mod._client is not None:
            try:
                client_mod._client.close()
            except Exception:
                pass
            client_mod._client = None


@pytest.fixture()
def mock_graph():
    """Activate respx router that intercepts all requests to graph.microsoft.com."""
    import respx
    with respx.mock(base_url="https://graph.microsoft.com/v1.0", assert_all_called=False) as router:
        yield router

