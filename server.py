import logging
import os
import sys

import outlook_mcp.tools  # noqa: F401 — registers all tools
from outlook_mcp.app import mcp

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
    stream=sys.stderr,
)
logger = logging.getLogger(__name__)

if __name__ == "__main__":
    if not os.environ.get("OUTLOOK_TOKEN"):
        logger.error("Required environment variable not set: OUTLOOK_TOKEN")
        sys.exit(1)

    port = os.environ.get("PORT")
    try:
        if port:
            logger.info("Starting MCP server on HTTP port %s", port)
            mcp.run(transport="streamable-http", host="0.0.0.0", port=int(port))
        else:
            logger.info("Starting MCP server on stdio")
            mcp.run(transport="stdio")
    except Exception:
        logger.exception("MCP server terminated with an unhandled exception")
        sys.exit(1)
