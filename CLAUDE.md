# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Commands

```bash
# Install dependencies
python3 -m venv .venv && source .venv/bin/activate && pip install -r requirements.txt

# Verify Outlook connectivity before touching Claude Desktop
python test_connection.py

# Run the MCP server manually (Claude Desktop does this automatically)
.venv/bin/python server.py
```

There are no automated tests — `test_connection.py` is the integration smoke test and requires Outlook to be running.

## Architecture

Three layers, each small enough to read in full:

```
Claude Desktop (stdio / MCP JSON-RPC)
    ↓
server.py          — FastMCP; registers 4 @mcp.tool() functions
    ↓
outlook_bridge.py  — subprocess wrapper; calls osascript, parses JSON response
    ↓
outlook_commands.js — JXA (JavaScript for Automation); talks to Outlook's scripting API
    ↓
Outlook for Mac (classic)
```

**Why JXA?** Python has no direct path to Outlook's scripting dictionary. The bridge spawns `osascript -l JavaScript outlook_commands.js <command> <args_json>` for every call; the JS side prints `{"ok": true, "data": ...}` or `{"ok": false, "error": ...}` to stdout.

**Key constraints:**
- Requires **classic** Outlook for Mac (not "New Outlook" — no scripting interface). macOS Automation permission must be granted to Terminal/Claude Desktop.
- `get_email` has no direct-by-ID lookup; it scans all folders linearly — keep email IDs short-lived.
- Email bodies are capped at 8 000 characters (`BODY_CAP` in `outlook_commands.js`) to limit context consumption.
- `search_emails(folder="all")` iterates every folder and is slow on large mailboxes.

## Adding a new tool

1. Add a function to `outlook_commands.js` that returns a plain JS object/array.
2. Register the function name in the `commands` map in `run()` (bottom of the file).
3. Add a `_run_jxa(...)` wrapper to `outlook_bridge.py`.
4. Decorate a new Python function with `@mcp.tool()` in `server.py`.
5. Restart Claude Desktop (Cmd+Q, reopen).

**Write capabilities:** The server is intentionally read-only. Adding send/delete/move tools changes the threat model — incoming emails could prompt-inject Claude into destructive actions. Any such addition warrants explicit deliberation.
