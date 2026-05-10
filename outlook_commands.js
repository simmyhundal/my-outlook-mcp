#!/usr/bin/osascript -l JavaScript
//
// outlook_commands.js
// ====================
// JXA (JavaScript for Automation) script that talks to Outlook for Mac.
// Called by outlook_bridge.py via `osascript -l JavaScript`.
// Receives a command name + JSON args, prints JSON result to stdout.
//
// IMPORTANT: read-only. No send, delete, move, or mark-as-read.
//

const BODY_CAP = 8000;

function safeGet(fn, fallback) {
    try { return fn(); } catch (_) { return fallback; }
}

// Recursively collect every mail folder across all accounts/sub-folders.
// Outlook.mailFolders() only returns top-level folders for the default account.
// For multi-account setups the second account's folders live as sub-folders
// under containers like "Sur mon ordinateur".
function getAllFolders() {
    const Outlook = Application("Microsoft Outlook");
    const result = [];
    function collect(folders) {
        for (const f of folders) {
            try {
                result.push(f);
                const subs = safeGet(() => f.mailFolders(), []);
                if (subs.length > 0) collect(subs);
            } catch (_) {}
        }
    }
    collect(safeGet(() => Outlook.mailFolders(), []));
    return result;
}

// Resolve a recipient object to {name, address}, handling both callable
// methods and plain string properties (JXA quirk with some Outlook versions).
function recipientToObj(r) {
    if (!r) return { name: "", address: "" };
    if (typeof r === "string") return { name: r, address: "" };
    return {
        name: typeof r.name === "function" ? safeGet(() => r.name(), "") : (r.name || ""),
        address: typeof r.address === "function" ? safeGet(() => r.address(), "") : (r.address || ""),
    };
}

function messageToObj(msg, includeBody) {
    const s = safeGet(() => msg.sender(), null);
    const obj = {
        id: safeGet(() => msg.id(), null),
        subject: safeGet(() => msg.subject(), ""),
        sender: !s ? "(unknown)" : typeof s === "string" ? s : (typeof s.name === "function" ? safeGet(() => s.name(), "(unknown)") : (s.name || "(unknown)")),
        senderEmail: !s || typeof s === "string" ? "" : (typeof s.address === "function" ? safeGet(() => s.address(), "") : (s.address || "")),
        received: safeGet(() => { const t = msg.timeReceived(); return t ? t.toISOString() : null; }, null),
        isRead: safeGet(() => msg.isRead(), null),
        hasAttachments: safeGet(() => msg.hasAttachment(), false),
    };
    if (includeBody) {
        let body = safeGet(() => msg.plainTextContent(), "") || "";
        if (body.length > BODY_CAP) {
            body = body.substring(0, BODY_CAP) + "\n\n[... truncated ...]";
        }
        obj.body = body;
    }
    return obj;
}

function list_emails(args) {
    const Outlook = Application("Microsoft Outlook");
    const folderName = args.folder || "Inbox";
    const limit = args.limit || 10;
    const unreadOnly = args.unread_only || false;

    let messages = [];
    if (folderName.toLowerCase() === "inbox") {
        messages = safeGet(() => Outlook.inbox.messages(), []);
    } else {
        // Collect from ALL folders matching the name (covers multiple accounts).
        const matched = getAllFolders().filter(f => safeGet(() => f.name(), "") === folderName);
        if (matched.length === 0) throw new Error(`Folder not found: ${folderName}`);
        for (const f of matched) {
            messages = messages.concat(safeGet(() => f.messages(), []));
        }
        // Sort newest-first across merged folders.
        messages.sort((a, b) => {
            const ta = safeGet(() => a.timeReceived(), null);
            const tb = safeGet(() => b.timeReceived(), null);
            if (!ta && !tb) return 0;
            if (!ta) return 1;
            if (!tb) return -1;
            return tb - ta;
        });
    }

    if (unreadOnly) messages = messages.filter(m => safeGet(() => !m.isRead(), false));
    // Deduplicate by id (multiple accounts can expose the same message).
    const seen = new Set();
    const unique = [];
    for (const m of messages) {
        const id = safeGet(() => m.id(), null);
        if (id !== null && !seen.has(id)) { seen.add(id); unique.push(m); }
    }
    return unique.slice(0, limit).map(m => messageToObj(m, false));
}

function get_email(args) {
    if (!args.id) throw new Error("Missing required argument: id");
    for (const folder of getAllFolders()) {
        const msgs = safeGet(() => folder.messages(), []);
        for (const msg of msgs) {
            if (safeGet(() => msg.id(), null) == args.id) {
                return messageToObj(msg, true);
            }
        }
    }
    throw new Error(`Email not found: ${args.id}`);
}

function search_emails(args) {
    const Outlook = Application("Microsoft Outlook");
    if (!args.query) throw new Error("Missing required argument: query");
    const q = args.query.toLowerCase();
    const limit = args.limit || 20;
    const folderName = args.folder || "Inbox";

    let foldersToSearch;
    if (folderName.toLowerCase() === "all") {
        foldersToSearch = getAllFolders();
    } else if (folderName.toLowerCase() === "inbox") {
        foldersToSearch = [Outlook.inbox];
    } else {
        // Match all folders with this name (multiple accounts).
        foldersToSearch = getAllFolders().filter(f => safeGet(() => f.name(), "") === folderName);
        if (foldersToSearch.length === 0) throw new Error(`Folder not found: ${folderName}`);
    }

    const matches = [];
    const seen = new Set();
    for (const folder of foldersToSearch) {
        const msgs = safeGet(() => folder.messages(), []);
        for (const m of msgs) {
            try {
                const id = safeGet(() => m.id(), null);
                if (id !== null && seen.has(id)) continue;
                const subj = safeGet(() => (m.subject() || "").toLowerCase(), "");
                const s = safeGet(() => m.sender(), null);
                const senderName = !s ? "" : typeof s === "string" ? s.toLowerCase() : (typeof s.name === "function" ? safeGet(() => s.name(), "").toLowerCase() : (s.name || "").toLowerCase());
                const senderEmail = !s || typeof s === "string" ? "" : (typeof s.address === "function" ? safeGet(() => s.address(), "").toLowerCase() : (s.address || "").toLowerCase());
                if (subj.includes(q) || senderName.includes(q) || senderEmail.includes(q)) {
                    if (id !== null) seen.add(id);
                    matches.push(messageToObj(m, false));
                    if (matches.length >= limit) return matches;
                }
            } catch (_) { continue; }
        }
    }
    return matches;
}

function list_folders() {
    const Outlook = Application("Microsoft Outlook");
    // Try to identify special folders by comparing object identity to
    // well-known Outlook properties (inbox, outbox, drafts, sent).
    // We compare name+unreadCount as a proxy since JXA object equality is unreliable.
    const specialCandidates = {};
    ["inbox", "outbox", "drafts"].forEach(prop => {
        try {
            const f = Outlook[prop];
            const key = safeGet(() => f.name() + "|" + f.unreadCount(), null);
            if (key) specialCandidates[key] = prop;
        } catch (_) {}
    });

    return getAllFolders()
        .map(f => {
            const name = safeGet(() => f.name(), null);
            const unreadCount = safeGet(() => f.unreadCount(), 0);
            const key = name + "|" + unreadCount;
            const obj = { name, unreadCount };
            if (specialCandidates[key]) obj.specialType = specialCandidates[key];
            // Heuristic: identify sent folders by French/English name patterns.
            if (name && /envoy|sent item/i.test(name)) obj.specialType = "sent";
            if (name && /brouillon|draft/i.test(name)) obj.specialType = obj.specialType || "drafts";
            if (name && /bo[iî]te de r[eé]ception|inbox/i.test(name)) obj.specialType = obj.specialType || "inbox";
            return obj;
        })
        .filter(f => f.name);
}

function run(argv) {
    try {
        const command = argv[0];
        const args = argv[1] ? JSON.parse(argv[1]) : {};
        const commands = { list_emails, get_email, search_emails, list_folders };
        const fn = commands[command];
        if (!fn) return JSON.stringify({ error: `Unknown command: ${command}` });
        const result = fn(args);
        return JSON.stringify({ ok: true, data: result });
    } catch (e) {
        return JSON.stringify({ ok: false, error: e.message || String(e) });
    }
}
