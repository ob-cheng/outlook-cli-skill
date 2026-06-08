# Feature Guide

This guide covers how the Outlook CLI actually works behind the scenes — so you know what to expect when you ask your agent to use it.

---

## Search & Email

### What gets searched
By default, search covers your **Inbox** and **Sent Items**. If you want to search other folders (Archive, a specific project folder, etc.), tell your agent to use `--folder "Folder Name"` — and you can use it multiple times to search several folders at once.

### How filtering works
When you ask your agent to "find emails from Alice", it uses `--filter-email`. This checks **From, To, and CC** — so it'll find emails *to* Alice as well as *from* her. If you want to narrow by domain (`--filter-domain company.com`), that works the same way.

There's a built-in guard: if an email has more than **20 recipients**, it gets skipped when you're using email filters. This prevents mass mailings and distribution lists from cluttering results.

### Auto-replies are automatically filtered out
Out of Office, Automatic reply, and similar auto-generated messages are silently excluded from search results — you'll never see OOO clutter unless you specifically look for it.

### Why message IDs can stop working
Outlook's internal message IDs (EntryIDs) can change when an email is moved between folders. If you saved a message ID and later it doesn't work, just ask your agent to re-search for it.

---

## Email Composition

### Drafts are the default
Every email your agent creates is saved as a **draft** — always. There's no way to accidentally send. The agent has to use `--send` *and* have direct sending enabled in config for actual sending to happen.

### If you want direct sending
Tell your agent: "enable direct sending for the outlook skill" — they'll run `python outlook.py config set send_mode send` and ask for confirmation before each send. Without it, everything stays in your Drafts folder for you to review and send manually.

### Attachments
Your agent can attach files to emails. Just provide the file path. Multiple attachments are supported. If a file doesn't exist, it's silently skipped rather than erroring out.

### Reply and forward behavior
When replying, your agent's response is added **above** the original message content (below a separator). Forwards work the same way — your added message goes on top.

### Custom drafting instructions
You can tell your agent how to write your emails. Just say something like:

> "Set my drafting instructions to: Keep it brief, bullet points only, always mention next steps."

The agent saves your exact words and follows them every time it drafts an email. To see or change them later:

```bash
python outlook.py config show
python outlook.py config set draft_instructions "New instructions here"
```

The agent never rewrites or rephrases your instructions — it uses them verbatim.

### Humanizer for natural writing
If you want your agent's drafts to sound less like an AI and more like you, enable the humanizer:

```bash
python outlook.py config set humanizer_enabled true
```

The agent will run a 29-pattern cleanup on every email body before saving the draft — stripping filler phrases, overused words, em dash spam, and other AI tells. If the humanizer skill isn't installed, the agent will ask if you want to install it.

To turn it off: `python outlook.py config set humanizer_enabled false`

### Draft review flow
Your agent creates a draft every time by default. No accidental sends. You can review everything in Outlook's Drafts folder before it goes out.

---

## Export

### What you get
The tool can export your emails in two formats:
- **Markdown** (default) — Obsidian-friendly files with YAML frontmatter, tags, and proper formatting. Each thread is grouped into one file with all messages in chronological order.
- **JSON** — structured data, best for AI processing or analysis.

### Thread grouping (smart default)
Emails on the same topic are automatically grouped into a single file, regardless of Re:/Fwd: prefixes. If you want each email as its own file, tell your agent to use `--no-threads`.

### Batch JSON mode
If you're exporting a lot of emails for AI processing, batch mode (`--format json --batch`) puts everything into one JSON file instead of one per thread. More efficient for token budgets.

### Incremental export (export since last time)
The tool can track what it's already exported and only grab new emails on subsequent runs. Here's how it works:

1. **First run** — exports everything matching your criteria, saves a timestamp
2. **Next run** — only exports what's arrived since that timestamp
3. **Reset** — delete the `extraction_state.json` file in the skill directory to start fresh

The state file lives in the skill's installation folder, **not** in your export directory — so you can export to different folders and the tracking still works.

### Content cleanup
When your agent exports an email, it:
- Converts HTML emails to clean markdown
- Strips out scripts, styles, and metadata
- Removes quoted reply text (the "--- From: ..." blocks)

The result is clean, readable content — no formatting debris.

---

## Calendar

### Which calendar gets used
Calendar commands search **all stores** accessible in Outlook. If you have multiple accounts (iCloud + Exchange, for example), all their calendars are searched — no configuration needed. The same applies to Tasks and Notes. Email search defaults to the primary store but accepts path-based folder targeting (e.g., `--folder "AccountName/Inbox"`) to reach any account.

### Creating events
Your agent can create meetings with attendees, locations, and descriptions. Invitations are sent automatically when attendees are included. Reminders default to 15 minutes.

---

## Tasks

### Status tracking
Tasks can be: `not_started`, `in_progress`, `completed`, `waiting`, or `deferred`. Ask your agent to check "my tasks" or "what's pending" to see them. Completed tasks are hidden by default — use `--all` to include them.

### Filtering
You can filter by priority (low/normal/high), due date range, status, or category. Ask naturally — "show me high priority tasks due this week."

---

## Notes

Quick sticky notes in Outlook. Create with body text, optionally color-code them (blue, green, pink, yellow, white — yellow is default), and assign categories for organization.

---

## Multi-Account Tips

If you have multiple email accounts set up in Outlook:

| What you're doing | Which account gets used |
|---|---|
| Searching/reading emails | Primary mailbox by default. Use `--folder "AccountName/Inbox"` to target another. |
| Calendar, Tasks, Notes | **All accounts.** Every store is searched — no switching needed. |

Email is the only command that doesn't auto-search across stores. Calendar, Tasks, and Notes search every store you have configured.
