from __future__ import annotations

import os

# Required — Microsoft Graph API bearer token
OUTLOOK_TOKEN: str = os.environ.get("OUTLOOK_TOKEN", "")

# Optional — HTTP proxy URL (e.g. for corporate networks)
OUTLOOK_PROXY: str = os.environ.get("OUTLOOK_PROXY", "")

# Optional — set to "true" to enable SSL certificate verification
VERIFY_SSL: bool = os.environ.get("OUTLOOK_VERIFY_SSL", "false").lower() == "true"

# Microsoft Graph API base URL — public endpoint, not company-specific
GRAPH_BASE_URL: str = "https://graph.microsoft.com/v1.0"
