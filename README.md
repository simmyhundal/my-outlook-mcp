# my-outlook-mcp

A local MCP server that gives Claude Desktop read access to Outlook for Mac — no OAuth, no Microsoft Graph, no cloud credentials. Built from scratch so every line of code that touches personal email is readable in under ten minutes.

## Impact

- **~10 emails/week** routed through Claude for CRM triage and follow-up tracking, keeping relationship context current without manual inbox archaeology
- Runs entirely on-device; zero data leaves the machine beyond what Claude Desktop already handles
- Handles a French-locale Outlook installation (real edge case, not hypothetical — see [AI Design Notes](#ai-design-notes))

## Project Overview

This server does two distinct jobs.

**Job 1 — Give Claude a local email interface.** Four read-only tools let Claude list folders, enumerate recent messages, fetch a full body, and run substring search. Claude picks the right call sequence autonomously; the human just asks in plain English.

**Job 2 — Enforce a hard write boundary by default.** The Outlook scripting API supports send, delete, and move — this server deliberately exposes none of them. The blast radius of a prompt-injection attack from a malicious incoming email is bounded to what can be *read*, not what can be *done*.

## Architecture

```
Claude Desktop
     │  stdio · MCP JSON-RPC
     ▼
server.py              FastMCP · registers 4 @mcp.tool() functions
     │
     ▼
outlook_bridge.py      subprocess wrapper · calls osascript · parses JSON
     │
     ▼
outlook_commands.js    JXA · talks to Outlook's scripting API
     │
     ▼
Outlook for Mac (classic)
```

| Layer | Language | Responsibility |
|---|---|---|
| `server.py` | Python | MCP protocol, tool registration, schema generation |
| `outlook_bridge.py` | Python | Subprocess management, error normalisation |
| `outlook_commands.js` | JXA (JavaScript for Automation) | Outlook scripting API, JSON serialisation |

**Why JXA?** Python has no direct path to Outlook's scripting dictionary. The bridge spawns `osascript -l JavaScript` per call — roughly 100 ms of overhead per tool invocation, acceptable for interactive use. The entire payload is a plain JSON string printed to stdout; no IPC framework, nothing to version.

## AI Design Notes

This server contains no LLM calls. The "AI layer" is Claude Desktop itself; this codebase is the tool surface Claude reasons over.

**Tool docstrings as prompts.** Each `@mcp.tool()` docstring is the only instruction Claude receives about when and how to call that function. They're written to prevent common failure modes: `list_folders()` tells Claude to look for `specialType='sent'` before touching sent mail; `search_emails` warns that `folder="all"` iterates every folder and is slow on large mailboxes. This is prompt engineering at the interface layer — not in a system prompt, but in the schema Claude reads before each tool call.

**Locale reliability.** Running on a French-locale macOS against a French-locale Outlook means folder names like `Éléments envoyés` instead of `Sent Items`. The JXA layer normalises folder discovery so Claude never needs to know the locale — it always receives English `specialType` tags regardless of what Outlook reports. This was a real production bug that broke multi-account folder discovery; fixing it required matching against both locale variants in the scripting layer.

**Read-only as a principle, not a gap.** Least-privilege by default: the scripting API supports write operations; the server does not expose them. This mirrors a standard posture for agentic systems — minimise the tool surface until there is a specific, deliberate reason to expand it. Adding write tools is a future option, not an oversight, and it warrants explicit reasoning about the expanded blast radius.

**Cost.** Zero incremental API cost beyond Claude Desktop's subscription. No embeddings, no background jobs, no telemetry.

## Key Invariants

- Outlook must be open and signed in; the server does not launch it
- Email IDs are positional, not stable GUIDs — do not persist them across sessions
- Bodies are capped at 8 000 characters per `read_email` call to limit context consumption
- `search_emails(folder="all")` scans every folder linearly — prefer a specific folder name when the target is known

## Requirements

- macOS
- **Classic** Outlook for Mac (not "New Outlook" — the new one has no scripting interface). Check via Outlook → About Microsoft Outlook; if it says "New Outlook," toggle it off in the Outlook menu.
- Python 3.10+
- Outlook open and signed in when the server runs

## Setup

```bash
cd ~/code/my-outlook-mcp
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

Verify the connection before wiring up to Claude:

```bash
python test_connection.py
```

On first run macOS will prompt: *"Terminal wants to control Microsoft Outlook."* Click **OK**, or grant it manually:

> System Settings → Privacy & Security → Automation → Terminal → Outlook ✓

You should see your folders and last 5 emails printed to stdout.

## Wire up to Claude Desktop

```bash
# Get the absolute paths you'll paste into config
echo "$(pwd)/.venv/bin/python"
echo "$(pwd)/server.py"
```

Open Claude Desktop → Settings → Developer → Edit Config:

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

Quit Claude Desktop completely (Cmd+Q) and reopen. `my-outlook` will appear in the connected tools list.

## Try it

```
What are my unread emails this week?
Find emails from [name] about the term sheet.
Summarize the last five messages from my advisor.
```

## Extending

To add a new tool:

1. Add a function to `outlook_commands.js` returning a plain JS object or array.
2. Register it in the `commands` map at the bottom of that file.
3. Add a `_run_jxa(...)` wrapper in `outlook_bridge.py`.
4. Decorate a new Python function with `@mcp.tool()` in `server.py`.
5. Restart Claude Desktop.

**Before adding write tools** (send, draft, move): the read-only constraint is load-bearing for the threat model. Adding write capability is a deliberate choice — at minimum, document the expanded blast radius and consider scoping each tool narrowly (e.g., draft-only, not send).
