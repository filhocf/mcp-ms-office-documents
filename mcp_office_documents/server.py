"""Entry point for mcp-ms-office-documents CLI."""

import os
import sys

_pkg_dir = os.path.dirname(os.path.abspath(__file__))
_project_root = os.path.dirname(_pkg_dir)

if os.path.exists(os.path.join(_project_root, "docx_tools")):
    if _project_root not in sys.path:
        sys.path.insert(0, _project_root)


def main():
    """Start the MCP Office Documents server.

    Supports:
    - stdio (default): for mcp.json integration
    - streamable-http: for standalone service (port 8958)

    Set MCP_TRANSPORT=streamable-http for HTTP mode.
    """
    from mcp_office_documents.app import mcp, config

    transport = os.environ.get("MCP_TRANSPORT", "stdio")

    if transport == "streamable-http":
        mcp.run(
            transport="streamable-http",
            host="0.0.0.0",
            port=int(os.environ.get("MCP_PORT", "8958")),
            log_level=config.logging.mcp_level_str,
            path="/mcp",
        )
    else:
        mcp.run(transport="stdio")


if __name__ == "__main__":
    main()
