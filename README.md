# my-outlook-mcp

A read-only MCP server that gives Claude Desktop access to your local
Outlook for Mac (classic) inbox. Built from scratch so you can read every
line of code that touches your email.

## What it does

Exposes four tools to Claude:

- `list_recent_emails(folder, limit, unread_only)` — list emails from a folder
- `read_email(email_id)` — fetch the full body of one email
- `search_emails(query, folder, limit)` — substring search by subject or sender
- `list_folders()` — list all mail folders with unread counts

## What it does NOT do

- Send email
- Delete email
- Move, mark-as-read, or modify anything
- Touch the network (no Microsoft Graph, no OAuth, no telemetry)
- Store credentials of any kind

If a malicious email tries to prompt-inject Claude into "deleting all my
mail," there is literally no tool here that can do that. The blast radius
is bounded by the code in this folder.

## Architecture

```
Claude Desktop
     │ stdio (MCP protocol — JSON-RPC)
     ▼
  server.py          ← FastMCP, registers the 4 tools
     │
     ▼
outlook_bridge.py    ← runs `osascript` as a subprocess
     │
     ▼
outlook_commands.js  ← JXA script, talks to Outlook's scripting API
     │
     ▼
  Outlook for Mac (classic)
```

Each layer is small enough to read in 5 minutes. That's the whole point.

## Requirements

- macOS
- **Classic** Outlook for Mac (not "New Outlook" — the new one has no
  scripting interface). Check via Outlook → About Microsoft Outlook. If
  it says "New Outlook," toggle it off in the Outlook menu.
- Python 3.10+
- Outlook must be open and signed in when the server runs.

## Setup

```bash
cd ~/code/my-outlook-mcp
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## Test before wiring up to Claude

Make sure Outlook is open, then:

```bash
python test_connection.py
```

The first run, macOS will pop up a dialog: "Terminal wants to control
Microsoft Outlook." Click **OK**. If you miss it, go to:

  System Settings → Privacy & Security → Automation → Terminal → enable Outlook

You should see your folders and last 5 emails print.

## Wire up to Claude Desktop

Get the absolute paths:

```bash
echo "$(pwd)/.venv/bin/python"
echo "$(pwd)/server.py"
```

Open Claude Desktop → Settings → Developer → Edit Config, and add:

```json
{
  "mcpServers": {
    "my-outlook": {
      "command": "/Users/YOU/code/my-outlook-mcp/.venv/bin/python",
      "args": ["/Users/YOU/code/my-outlook-mcp/server.py"]
    }
  }
}
```

Quit Claude Desktop completely (Cmd+Q) and reopen. You should see
`my-outlook` in the connected tools list.

## Try it

In a new Claude conversation:

> What are my unread emails today?
>
> Find emails from my advisor about the thesis proposal.
>
> Summarize the most recent email in my inbox.

## Adding tools later

If you want to extend this — for example, list calendar events — the
pattern is:

1. Add a function to `outlook_commands.js` that returns a JS object/array.
2. Register it in the `commands` map at the bottom of that file.
3. Add a wrapper to `outlook_bridge.py` that calls `_run_jxa(...)`.
4. Add a `@mcp.tool()`-decorated function to `server.py`.
5. Restart Claude Desktop.

If you want to add **write** capabilities (send, draft, move), do it
deliberately. The reason this server is safe is precisely that those
tools don't exist. Adding them shifts the threat model — at minimum,
think about prompt injection from incoming email.
