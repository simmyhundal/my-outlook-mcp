#!/usr/bin/osascript -l JavaScript
//
// outlook_commands.js
// ====================
// This is a JXA (JavaScript for Automation) script that talks to the Outlook
// app on your Mac. It's the same mechanism AppleScript uses — macOS lets
// scripts control native apps, and Outlook for Mac (classic) exposes a
// scripting interface we can drive.
//
// The Python server (server.py) calls this script via `osascript`, passing
// a command name and a JSON blob of arguments. We dispatch to the right
// function and print JSON back to stdout. The Python side parses the JSON.
//
// Why JXA and not pure Python? Python can't talk to Outlook's scripting
// dictionary directly. JXA is the bridge. (You could also use AppleScript,
// but JXA gives us real JSON.stringify which is much cleaner.)
//
// IMPORTANT: this script is read-only. There is no `send`, `delete`,
// `move`, or `mark as read` function. If something goes wrong, the worst
// it can do is read mail you can already see in Outlook.
//

// Cap on how much email body text we return per message. Long emails can
// blow out Claude's context window. Bump this if you need more, but be
// aware it costs tokens.
const BODY_CAP = 8000;

// Helper: turn an Outlook message object into a plain JS object we can
// stringify. We deliberately pull only the fields we need — never anything
// like attachments-as-bytes that could cause us to ship gigabytes back.
function messageToObj(msg, includeBody) {
    const obj = {
        id: msg.id(),
        subject: msg.subject(),
        sender: msg.sender() ? msg.sender().name() : "(unknown)",
        senderEmail: msg.sender() ? msg.sender().address() : "",
        // .timeReceived() returns a JS Date; toISOString gives a stable format
        received: msg.timeReceived() ? msg.timeReceived().toISOString() : null,
        isRead: msg.isRead(),
        hasAttachments: msg.hasAttachment(),
    };
    if (includeBody) {
        // plainTextContent strips the HTML for us. If you ever want HTML,
        // use msg.content() instead — but be careful, it's much bigger.
        let body = msg.plainTextContent() || "";
        if (body.length > BODY_CAP) {
            body = body.substring(0, BODY_CAP) + "\n\n[... truncated ...]";
        }
        obj.body = body;
    }
    return obj;
}

// === Tool implementations ===
// Each function here is callable from server.py. They all return plain
// JS objects/arrays that get JSON.stringified at the end.

// List recent emails from a folder. Default: inbox, 10 most recent.
function list_emails(args) {
    const Outlook = Application("Microsoft Outlook");
    const folderName = args.folder || "Inbox";
    const limit = args.limit || 10;
    const unreadOnly = args.unread_only || false;

    // Outlook's scripting API exposes folders by name. "Inbox" is the
    // default mail folder; you can also pass "Sent Items", "Drafts", etc.
    let folder;
    if (folderName.toLowerCase() === "inbox") {
        folder = Outlook.inbox;
    } else {
        // Look up by exact name across all mail folders.
        const allFolders = Outlook.mailFolders();
        folder = allFolders.find(f => f.name() === folderName);
        if (!folder) {
            throw new Error(`Folder not found: ${folderName}`);
        }
    }

    // Pull messages. Outlook returns them newest-first in the inbox.
    let messages = folder.messages();
    if (unreadOnly) {
        messages = messages.filter(m => !m.isRead());
    }
    messages = messages.slice(0, limit);

    return messages.map(m => messageToObj(m, false));
}

// Get the full body of one email by ID.
function get_email(args) {
    const Outlook = Application("Microsoft Outlook");
    if (!args.id) throw new Error("Missing required argument: id");

    // We have to search for the message by ID. Outlook's JXA bindings
    // don't have a direct getById, so we iterate. This is fine for normal
    // use — you'll have looked it up via list_emails first, so it's still
    // in cache.
    const folders = Outlook.mailFolders();
    for (const folder of folders) {
        const msgs = folder.messages();
        for (const msg of msgs) {
            if (msg.id() == args.id) {
                return messageToObj(msg, true);
            }
        }
    }
    throw new Error(`Email not found: ${args.id}`);
}

// Search emails by subject/sender substring.
function search_emails(args) {
    const Outlook = Application("Microsoft Outlook");
    if (!args.query) throw new Error("Missing required argument: query");
    const q = args.query.toLowerCase();
    const limit = args.limit || 20;
    const folderName = args.folder || "Inbox";

    let foldersToSearch;
    if (folderName.toLowerCase() === "all") {
        // "all" iterates every folder. Slow on big mailboxes — only use
        // when you can't predict where the message is.
        foldersToSearch = Outlook.mailFolders();
    } else if (folderName.toLowerCase() === "inbox") {
        foldersToSearch = [Outlook.inbox];
    } else {
        const folder = Outlook.mailFolders().find(f => f.name() === folderName);
        if (!folder) throw new Error(`Folder not found: ${folderName}`);
        foldersToSearch = [folder];
    }

    const matches = [];
    for (const folder of foldersToSearch) {
        const msgs = folder.messages();
        for (const m of msgs) {
            const subj = (m.subject() || "").toLowerCase();
            const senderName = m.sender() ? (m.sender().name() || "").toLowerCase() : "";
            const senderEmail = m.sender() ? (m.sender().address() || "").toLowerCase() : "";
            if (subj.includes(q) || senderName.includes(q) || senderEmail.includes(q)) {
                matches.push(messageToObj(m, false));
                if (matches.length >= limit) return matches;
            }
        }
    }
    return matches;
}

// List the names of all mail folders. Useful so you know what to pass
// to list_emails / search_emails.
function list_folders() {
    const Outlook = Application("Microsoft Outlook");
    return Outlook.mailFolders().map(f => ({
        name: f.name(),
        unreadCount: f.unreadCount(),
    }));
}

// === Dispatcher ===
// osascript passes argv to the run() function. We expect:
//   argv[0] = command name
//   argv[1] = JSON string of arguments (optional)
//
function run(argv) {
    try {
        const command = argv[0];
        const args = argv[1] ? JSON.parse(argv[1]) : {};

        const commands = {
            list_emails,
            get_email,
            search_emails,
            list_folders,
        };

        const fn = commands[command];
        if (!fn) {
            return JSON.stringify({ error: `Unknown command: ${command}` });
        }

        const result = fn(args);
        return JSON.stringify({ ok: true, data: result });
    } catch (e) {
        return JSON.stringify({ ok: false, error: e.message || String(e) });
    }
}
