"""
mcp_server.py — Root Facade for Model Context Protocol (MCP) Server.

Thin Facade pointing to server.mcp_server for strict Layered Modular Architecture.
"""
from __future__ import annotations

from server.mcp_server import (
    main,
    handle_rpc_request,
    run_stdio_server,
    TOOL_DEFINITIONS,
    TOOL_DISPATCHER,
    SERVER_NAME,
    SERVER_VERSION,
)

__all__ = [
    "main",
    "handle_rpc_request",
    "run_stdio_server",
    "TOOL_DEFINITIONS",
    "TOOL_DISPATCHER",
    "SERVER_NAME",
    "SERVER_VERSION",
]

if __name__ == "__main__":
    main()
