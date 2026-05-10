"""
server.py
=========

The MCP server. This is what Claude Desktop launches and talks to over stdio.

We use FastMCP (from the official `mcp` Python SDK), which lets us define
tools as plain Python functions decorated with @mcp.tool(). FastMCP handles
all the JSON-RPC plumbing automatically — parameter validation, schema
generation, error formatting.

How Claude Desktop uses this:
  1. You add this server to claude_desktop_config.json.
  2. When Claude starts, it spawns this script as a subprocess.
  3. Claude communicates over stdin/stdout using the MCP protocol.
  4. When you ask Claude to read your email, it picks the right tool and
     calls it. FastMCP runs the corresponding function, our function calls
     into outlook_bridge.py, which runs the JXA script, which talks to
     Outlook.

Read-only by design: there are NO send/delete/modify tools here, even
though the Outlook scripting API supports them. Adding write capability
is a deliberate choice you'd make later, with eyes open.
"""

from mcp.server.fastmcp import FastMCP

import outlook_bridge

# Server name shows up in Claude's UI when you list connected MCPs.
mcp = FastMCP("my-outlook")


@mcp.tool()
def list_recent_emails(folder: str = "Inbox", limit: int = 10, unread_only: bool = False) -> list:
    """List recent emails from a folder, newest first.

    Args:
        folder: Folder name to read from. Default "Inbox". Use list_folders() to discover names.
        limit: Maximum number of emails to return (default 10).
        unread_only: If true, only return unread emails.

    Returns a list of email summaries (without body text). Use read_email
    with the id to fetch the full body of one.
    """
    return outlook_bridge.list_emails(folder=folder, limit=limit, unread_only=unread_only)


@mcp.tool()
def read_email(email_id: str) -> dict:
    """Fetch the full content of one email by ID.

    Args:
        email_id: The id field from a list_recent_emails or search_emails result.

    Returns the email including subject, sender, received time, and the
    plain-text body (truncated to 8000 characters).
    """
    return outlook_bridge.get_email(email_id)


@mcp.tool()
def search_emails(query: str, folder: str = "Inbox", limit: int = 20) -> list:
    """Search for emails by subject, sender, or recipient (case-insensitive substring match).

    Args:
        query: Substring to look for in subject, sender name, sender email, recipient name, or recipient email.
        folder: Folder to search in. Use "all" to search every folder (slow on large mailboxes).
              Use the exact French folder name (e.g. "Éléments envoyés") for Sent Items.
        limit: Max matches to return.

    Returns a list of matching email summaries.
    """
    return outlook_bridge.search_emails(query=query, folder=folder, limit=limit)


@mcp.tool()
def list_folders() -> list:
    """List all mail folders with their unread counts.

    Useful before calling list_recent_emails on a non-Inbox folder, to find
    the exact folder name.
    """
    return outlook_bridge.list_folders()


if __name__ == "__main__":
    # stdio transport is what Claude Desktop expects. Don't change this
    # unless you're hooking the server up to something else.
    mcp.run(transport="stdio")
