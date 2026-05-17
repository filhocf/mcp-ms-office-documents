"""Entry point for mcp-ms-office-documents CLI."""

import os
import sys

# When installed as package, modules like docx_tools/ are in site-packages.
# When running from repo, they're in the project root.
# We need to ensure the project root is in path for both cases.
_pkg_dir = os.path.dirname(os.path.abspath(__file__))
_project_root = os.path.dirname(_pkg_dir)

# Add project root for repo-local runs
if os.path.exists(os.path.join(_project_root, "docx_tools")):
    if _project_root not in sys.path:
        sys.path.insert(0, _project_root)


def main():
    """Start the MCP Office Documents server."""
    # Import here to ensure path is set
    from mcp_office_documents.app import mcp, config

    mcp.run(
        transport="streamable-http",
        host="0.0.0.0",
        port=8958,
        log_level=config.logging.mcp_level_str,
        path="/mcp",
    )


if __name__ == "__main__":
    main()
