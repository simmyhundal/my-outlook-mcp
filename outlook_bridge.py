"""
outlook_bridge.py
=================

Thin Python wrapper around the JXA script (outlook_commands.js).
Each function here:
  1. Spawns `osascript -l JavaScript outlook_commands.js <command> <args_json>`
  2. Captures stdout
  3. Parses the JSON response
  4. Raises an exception if the JXA side reported an error

We keep this layer dumb on purpose. All the Outlook logic lives in JS;
this just shuttles bytes back and forth. That separation makes it easy
to read the JS file and verify exactly what we're doing to your mailbox.
"""

import json
import subprocess
from pathlib import Path

# Path to the JXA script — sits next to this file.
SCRIPT_PATH = Path(__file__).parent / "outlook_commands.js"


def _run_jxa(command: str, args: dict | None = None) -> dict:
    """Execute a command in the JXA script and return the parsed result.

    Raises RuntimeError if osascript fails or the JXA side reports an error.
    """
    args_json = json.dumps(args or {})

    # We invoke osascript directly. -l JavaScript tells it to interpret
    # our file as JXA rather than AppleScript.
    proc = subprocess.run(
        [
            "osascript",
            "-l", "JavaScript",
            str(SCRIPT_PATH),
            command,
            args_json,
        ],
        capture_output=True,
        text=True,
        # 60s is generous; most commands return in well under a second.
        # Search across all folders on a big mailbox could be slower.
        timeout=60,
    )

    if proc.returncode != 0:
        # osascript itself failed — usually means Outlook isn't running
        # or we don't have automation permission.
        raise RuntimeError(
            f"osascript failed (exit {proc.returncode}): {proc.stderr.strip()}"
        )

    try:
        result = json.loads(proc.stdout)
    except json.JSONDecodeError as e:
        raise RuntimeError(
            f"Could not parse JXA output as JSON: {e}\nRaw output: {proc.stdout[:500]}"
        )

    if not result.get("ok"):
        # The JXA dispatcher caught an exception; propagate it.
        raise RuntimeError(f"Outlook command failed: {result.get('error')}")

    return result["data"]


# === Public API — one function per JXA command ===

def list_emails(folder: str = "Inbox", limit: int = 10, unread_only: bool = False) -> list:
    return _run_jxa("list_emails", {
        "folder": folder,
        "limit": limit,
        "unread_only": unread_only,
    })


def get_email(email_id: str) -> dict:
    return _run_jxa("get_email", {"id": email_id})


def search_emails(query: str, folder: str = "Inbox", limit: int = 20) -> list:
    return _run_jxa("search_emails", {
        "query": query,
        "folder": folder,
        "limit": limit,
    })


def list_folders() -> list:
    return _run_jxa("list_folders", {})
