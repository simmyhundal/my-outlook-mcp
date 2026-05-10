"""
test_connection.py
==================

Run this BEFORE wiring the server up to Claude Desktop. It exercises the
JXA bridge directly so you can verify:

  1. Outlook is running and you have classic Outlook (not "New Outlook")
  2. macOS Automation permission is granted
  3. The JXA script returns data we can parse

If this prints emails, the MCP server will work. If it errors, fix the
problem here first — debugging through Claude is a pain.

Usage:
    python test_connection.py
"""

import outlook_bridge

print("Testing Outlook bridge...\n")

print("=== Mail folders ===")
folders = outlook_bridge.list_folders()
for f in folders[:10]:
    print(f"  {f['name']:30}  unread: {f['unreadCount']}")

print("\n=== Last 5 inbox emails ===")
emails = outlook_bridge.list_emails(folder="Inbox", limit=5)
for e in emails:
    star = "*" if not e["isRead"] else " "
    print(f"  {star} [{e['received']}] {e['sender']}: {e['subject']}")

if emails:
    print(f"\n=== Full body of first email (id={emails[0]['id']}) ===")
    full = outlook_bridge.get_email(emails[0]["id"])
    body = full.get("body", "")
    print(body[:500] + ("..." if len(body) > 500 else ""))

print("\nAll good. You can now wire this up to Claude Desktop.")
