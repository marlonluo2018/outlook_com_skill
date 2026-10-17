---
name: outlook_com_skill
description: Microsoft Outlook email management - search, list, compose, reply, forward, download attachments, thread tracking
triggers: [
  "check email", "check inbox", "any new emails", "what's new",
  "show recent emails", "show emails", "list emails",
  "find emails about", "find all emails from", "search for emails",
  "find thread", "find conversation",
  "find related",
  "draft email", "compose", "write email", "new email",
  "reply", "forward", "redirect", "send to",
  "batch forward", "mass forward", "forward to multiple",
  "get email", "view email", "show email details",
  "download attachment", "save attachment", "get attachment",
  "lookup contact",
  "use as template", "edit email html", "modify email", "create from template",
  "get signature", "view signature", "update signature", "change signature",
  "calendar", "show calendar", "my schedule", "what meetings",
  "accept meeting", "decline meeting", "tentative", "propose new time",
  "check availability", "check calendar", "free busy", "freebusy", "is available", "check schedule", "find time", "others calendar", "colleague calendar",
]
operations: ["find-recent", "find", "compose", "reply", "forward", "redirect", "batch-forward", "download-attachment", "contact-lookup", "find-thread", "find-related", "get-email", "recall", "get-html", "edit-html", "send-draft", "get-signature", "update-signature", "list-calendar", "respond-meeting", "check-availability"]
---

# Outlook Skill

> **⚠️ ALL emails use HTML format:** `<br>` for line breaks (avoid stacked `<p>` — Outlook adds large gaps), `<strong>bold</strong>`
> **Send commands send immediately.** Run them only after the user approved the draft (Send Flow in `assistant_brain/workflows/EMAIL_WORKFLOW.md`).
> **⚠️ No closing or signature in email body** — Outlook auto-appends signature. Do not add "Thanks, Marlon" or similar.
> **Codex execution:** Outlook COM requires the interactive Outlook desktop session. In Codex, run Outlook-backed commands with desktop/elevated execution directly; do not try the background sandbox first.

## Terminology

| Term | Meaning | Commands |
|------|---------|----------|
| **Draft review** | AI shows email text in chat for user approval before sending | `reply`, `compose`, `forward` |
| **Template editing** | Use an existing email as template → creates new item in Outlook Drafts folder for preview/iteration | `edit-html` → `send-draft` |

- When user says **"draft"**, **"先draft"**, or **"let me review"** → they want **draft review** (show in chat, wait for approval, then send via standard commands).
- **Template editing** is a separate workflow for modifying existing emails as templates. It saves to Outlook Drafts folder — only use when user explicitly asks to "edit as template" or "use as template".

## Commands

### Find Recent Emails
```bash
py -3 scripts/outlook_skill.py find-recent --days 7
py -3 scripts/outlook_skill.py find-recent --days 7 --json
```
- Default: **Inbox + Sent Items** (searches both folders to capture full context)
- Shows: To/CC, attachments, body preview, folder indicator (📥/📤)
- `--days`: 1-365 (default: 7)
- `--folder`: override to search a single folder only
- `--json`: stable machine-readable output including both `entry_id` and `conversation_id`

### Find Emails
```bash
py -3 scripts/outlook_skill.py find --from "Beng" --days 14
py -3 scripts/outlook_skill.py find --subject "voucher" --days 7 --limit 5
py -3 scripts/outlook_skill.py find --to "carmina" --days 30
py -3 scripts/outlook_skill.py find --body "approved" --days 7
py -3 scripts/outlook_skill.py find --from "beng" --subject "voucher" --days 14
```
- `--from`: search by sender name/email
- `--subject`: search by subject
- `--to`: search by recipient name/email
- `--body`: search by body content
- `--days`: 1-365 (default: 14)
- `--limit`: max results (default: all)
- `--folder`: single folder to search
- `--folders`: comma-separated folder names for cross-folder search
- Multiple fields can be combined (AND logic): `--from "X" --subject "Y"` returns emails matching both
- Default folder depends on field:
  - `--subject`, `--body` → **Inbox + Sent Items** (both)
  - `--from` → **Inbox** only
  - `--to` → **Sent Items** only
- `--to` matches recipients in sent mail using **To + CC** fields and resolved Outlook names/addresses
- **AI guidance:** start with a small recent window first (usually 7-14 days)
- If the first search does not find the email, widen the date range gradually and make the query more specific before broadening further
- Use `find-thread` or `find-related` when older or broader history is needed

### Find Thread
```bash
py -3 scripts/outlook_skill.py find-thread "<email_id>"
py -3 scripts/outlook_skill.py find-thread "<email_id>" --fuzzy
py -3 scripts/outlook_skill.py find-thread "<email_id>" --brief
```
- **Auto-searches Inbox + Sent Items** — thread completeness requires both
- Finds ALL emails sharing the same ConversationID
- Subjects can differ (RE:/Fwd: prefixes, topic changes don't matter)
- Uses Restrict first for speed, falls back to full-folder scan when needed
- `--fuzzy`: Also find emails with similar subjects (token overlap ≥ 0.6) when ConversationID breaks
- `--brief`: Compact single-line output (still shows email ID)
- Results sorted chronologically (oldest first)
- Shows thread summary: message count, participants, date span

### Find Related Emails
```bash
py -3 scripts/outlook_skill.py find-related "<email_id>"
py -3 scripts/outlook_skill.py find-related "<email_id>" --exclude-thread
py -3 scripts/outlook_skill.py find-related "<email_id>" --strategies recipient,keyword
py -3 scripts/outlook_skill.py find-related "<email_id>" --limit 10 --brief
```
- **Auto-searches Inbox + Sent Items** — single merged scan per folder (fast)
- Output includes relevance stars (★) and strategy name per result
- Multi-strategy search for emails related to a given email:
  - **thread** (★5): Same ConversationID
  - **sender** (★3-4): Same sender within time window + topic keyword overlap
  - **recipient** (★3): Shared recipients (≥2 people overlap in To/CC)
  - **keyword** (★2-3): Shared meaningful topic keywords from subject + body
- Generic noise terms such as external/training/request are intentionally ignored during keyword extraction
- Sender and keyword matching are intentionally tighter to reduce unrelated same-sender and boilerplate matches
- Results sorted by confidence then time (newest first within same confidence)
- `--strategies`: comma-separated (default: all four)
- `--exclude-thread`: skip thread strategy (useful after `find-thread` to avoid duplicates)
- `--limit N`: Limit results returned (default: 20, configurable in config.py)
- `--brief`: Compact single-line output (still shows email ID)

### Contact Lookup (Use Before Search by Email)
```bash
py -3 scripts/outlook_skill.py lookup-contact "user@domain.com"
py -3 scripts/outlook_skill.py lookup-contact "HONG YANG"
```
- Accepts: Email address OR display name
- Auto-detects: `@` present → email lookup; no `@` → name lookup via Exchange GAL
- Returns: Display name, email, alias, company, department, job title, office, phone, mobile, location
- **Why:** Outlook search by email address unreliable; use display name instead
- **Name lookup:** Resolves display names against Exchange Global Address List

### Sending Strategy (AI auto-selects — never ask user)

`reply` cannot remove recipients — its `--to`/`--cc` only append. Compare the source To/CC (from `get-email`) with the intended list and pick:

| Need | Command |
| ---- | ------- |
| Same/more recipients | `reply` |
| Sender only | `reply --only` |
| Remove or replace some recipients, keep the thread | `redirect` |
| New audience that needs thread context | `forward` |
| No thread context needed | `compose` with `Re:` subject |

### Body Transport (all send commands)

Write the HTML body with the Write tool to a new, uniquely named file in `outbox/`, then pass `--body-file`. Never put body text inside a shell command string — Bash/PowerShell expand `$60` etc. before the script sees it. After a successful send the CLI deletes the file (only files inside `outbox/`); on failure it is kept for retry. Use `<br>` for line breaks rather than `<p>` blocks.

### Reply

```bash
py -3 assistant_brain/skills/outlook_com_skill/scripts/outlook_skill.py reply "<email_id>" --body-file "outbox/T127-reply-20260924.html" --cc "extra@ibm.com" --attach "C:/path/file.pdf"
```

- **Default: reply-all** — keeps ALL original To + CC recipients. `--to`/`--cc` APPEND to existing.
- **`--only`: reply to From (sender) only** — AI selects this when the response is for the sender alone.
- `--attach`: File path(s) to attach (comma separated for multiple).
- `--importance`: Set priority flag (`high` or `low`) — shows ❗ in recipient's inbox.
- `reply` logs progress steps to stderr for timeout diagnosis.
- `reply` prints the Sent Items EntryID by default after sending.

### Compose Email

```bash
py -3 assistant_brain/skills/outlook_com_skill/scripts/outlook_skill.py compose --to "email@ibm.com" --subject "text" --body-file "outbox/<name>.html" --attach "C:/path/file.pdf"
```

- `--attach`: File path(s) to attach (comma separated for multiple)
- `--inline-image`: Embed image inline via CID (format: `filepath:cid_name`, comma separated)
- `--importance`: Set priority flag (`high` or `low`) — shows ❗ in recipient's inbox
- Sends immediately when called

### Forward (single)

```bash
py -3 assistant_brain/skills/outlook_com_skill/scripts/outlook_skill.py forward "<email_id>" --to "user1@ibm.com" --cc "manager@ibm.com" --body-file "outbox/<name>.html"
py -3 assistant_brain/skills/outlook_com_skill/scripts/outlook_skill.py forward "<email_id>" --to "user@domain.com" --attach "C:/path/file.pdf"
```

- Forwards an email to specified recipients
- `--to` (required): Comma-separated list of To recipients
- `--cc` (optional): Comma-separated list of CC recipients
- `--body-file` (optional): custom HTML message to prepend (see Body Transport)
- `--attach` (optional): File path(s) to attach (comma separated for multiple)
- Subject auto-prefixed with `FW:`
- Preserves original email formatting

### Batch Forward

```bash
py -3 assistant_brain/skills/outlook_com_skill/scripts/outlook_skill.py batch-forward "<email_id>" "recipients.csv" --body-file "outbox/<name>.html"
```

- CSV: single column named "email" (supports BOM encoding)
- `--body-file` (optional): custom HTML message to prepend (see Body Transport)
- Uses BCC for privacy
- Preserves original email formatting
- Automatically splits large recipient lists into batches
- **Batch size:** Configured in [`backend/config.py`](backend/config.py) (default: 500)

### Redirect (Clear Recipients + New TO/CC)

```bash
py -3 assistant_brain/skills/outlook_com_skill/scripts/outlook_skill.py redirect "<email_id>" --to "a@b.com" --cc "b@b.com" --body-file "outbox/<name>.html"
```

- Clears all existing TO and CC recipients, then adds new ones
- Preserves original email body as quoted content (like forward)
- `--body-file`: HTML message prepended above original content (see Body Transport)
- `--to` (required): New TO recipients (comma separated)
- `--cc`: New CC recipients (comma separated)
- `--attach`: File path(s) to attach (comma separated)
- Use when inherited recipients must be removed or replaced while keeping the thread

### Recall Sent Email
```bash
py -3 scripts/outlook_skill.py recall "<email_id>"
```
- Recalls a sent email via Exchange server (removes from recipients' inbox)
- Email must be in Sent Items (`email_item.Sent == True`)
- **Limitations:**
  - Only works for recipients on the same Exchange/Microsoft 365 organization
  - Only works if recipient hasn't read the message yet
  - External recipients (different mail server) will still see the email
- You will receive a "Message Recall Report" email from Exchange indicating success/failure per recipient

### Download Attachment
```bash
py -3 scripts/outlook_skill.py download-attachment "<email_id>"
py -3 scripts/outlook_skill.py download-attachment "<email_id>" --output-dir "C:\temp"
py -3 scripts/outlook_skill.py download-attachment "<email_id>" --filename "report.pdf"
```
- Saves email attachments to local directory (default: `~/Downloads`)
- `--output-dir`: Override save location
- `--filename`: Download only a specific attachment by name
- `--all`: Include embedded images (default: skips inline images, keeps PDFs/docs)
- Use with `reply --attach` or `forward --attach` to relay attachments

### Get Full Email Details
```bash
py -3 scripts/outlook_skill.py get-email "<email_id_1>" "<email_id_2>" ...
py -3 scripts/outlook_skill.py get-email "<email_id>" --truncate 1000
```
- Returns complete email details: full body, all attachments, metadata.
- **Batch retrieval:** Supports passing multiple email IDs separated by spaces to print them sequentially in one run.
- `--truncate N`: Truncate the printed body text to `N` characters to save terminal output and optimize context window/tokens.
- Use after search/thread/related to read the actual content.
- **Embedded images:** Auto-extracted to `<repo>/downloads/outlook_inline/<id>\` (gitignored; reused on re-read, folders older than 14 days pruned automatically). Paths printed in output — use Read tool to view. Never stored in the OS temp dir.

### Get Email HTML (Template Editing — Step 1: Read)
```bash
py -3 scripts/outlook_skill.py get-html "<email_id>"
```
- Returns the **raw HTMLBody** of an email (not plain text)
- Use this to inspect the HTML structure before editing
- Output wrapped in `HTML_START` / `HTML_END` markers for easy parsing
- Shows subject, sender, recipients, and HTML character count

### Edit Email HTML (Template Editing)
```bash
py -3 scripts/outlook_skill.py edit-html "<email_id>" --replace "old text::new text"
py -3 scripts/outlook_skill.py edit-html "<email_id>" --replace "Q1 2025::Q2 2025" --replace "January::April"
py -3 scripts/outlook_skill.py edit-html "<email_id>" --replace "old::new" --subject "New Subject" --to "new@ibm.com"
py -3 scripts/outlook_skill.py edit-html "<email_id>" --body-file "C:\temp\modified.html"
py -3 scripts/outlook_skill.py edit-html "<email_id>" --replace "old::new" --copy-attachments
```
- Takes a source email as template, applies modifications, saves as **new draft**
- Original email is **never modified** — always creates a fresh draft
- Result appears in your **Drafts folder** in Outlook
- `--replace "old::new"`: Text/HTML replacement (repeatable, uses `::` separator)
- `--subject`: Override subject line
- `--to`: Override To recipients (comma separated)
- `--cc`: Override CC recipients (comma separated)
- `--body-file`: Replace entire HTML body from a local file
- `--copy-attachments`: Copy file attachments from source (skips embedded images)
- Inline/embedded images (CID-referenced) are **always preserved** automatically
- **AI workflow:** `get-html` → analyze → `edit-html` with targeted replacements

### Send Draft
```bash
py -3 scripts/outlook_skill.py send-draft "<draft_email_id>"
```
- Sends an existing draft email from the Drafts folder
- Validates: must be unsent, must have at least one recipient
- **AI workflow:** `edit-html` (iterate until user approves) → `send-draft`

### Get Signature
```bash
py -3 scripts/outlook_skill.py get-signature
py -3 scripts/outlook_skill.py get-signature "SignatureName"
py -3 scripts/outlook_skill.py get-signature "SignatureName" --format html
```
- No args: list all signatures with plain text content
- With name: show that signature's content
- `--format html`: show raw HTML instead of plain text

### Update Signature
```bash
py -3 scripts/outlook_skill.py update-signature "SignatureName" --text "Line 1\nLine 2\nLine 3"
py -3 scripts/outlook_skill.py update-signature "SignatureName" --find "old text" --replace "new text"
py -3 scripts/outlook_skill.py update-signature "SignatureName" --after "anchor text" --insert "new line 1\nnew line 2"
py -3 scripts/outlook_skill.py update-signature "SignatureName" --body "<p>full HTML</p>"
```
- `--text`: full plain text replacement (use `\n` for line breaks) — **preferred mode**
- `--find`/`--replace`: targeted single-line substitution (e.g., update OOO dates)
- `--after`/`--insert`: insert new lines after a matched text (use `\n` for multiple lines)
- `--body`: full raw HTML replacement
- Updates all 3 signature files: `.htm`, `.txt`, `.rtf`
- Creates `.bak` backup before modifying
- **⚠️ Show current signature to user first, confirm changes before applying**
- **⚠️ Outlook must be restarted for changes to take effect:**
```bash
powershell -Command "Get-Process outlook -ErrorAction SilentlyContinue | Stop-Process -Force; Start-Sleep -Seconds 2; Start-Process outlook"
```

### List Calendar
```bash
py -3 scripts/outlook_skill.py list-calendar                             # today
py -3 scripts/outlook_skill.py list-calendar --days 7                    # next 7 days
py -3 scripts/outlook_skill.py list-calendar --start "2026-06-23" --end "2026-06-30"
py -3 scripts/outlook_skill.py list-calendar --user "xuhuanxh@cn.ibm.com" --date "2026-09-28"
```
- Shows appointments grouped by date: time range, subject, location, organizer, response status
- `--days`: Number of days ahead (default: 1 = today only)
- `--start`/`--end`: Explicit date range (YYYY-MM-DD)
- `--user` / `--contact`: Check calendar/availability for another colleague
- Includes recurring meetings (IncludeRecurrences enabled)
- Response status: Organized, Accepted, Tentative, Declined, Not Responded

### Check Availability / Free-Busy (Others' Calendar)
```bash
# Check availability for a single colleague
py -3 scripts/outlook_skill.py check-availability "user@ibm.com"

# Check multiple attendees + find common free slots (with your own calendar)
py -3 scripts/outlook_skill.py check-availability "user1@ibm.com, user2@ibm.com" --date "2026-09-28" --compare

# Show ONLY the common free slots for a 60-minute meeting
py -3 scripts/outlook_skill.py check-availability "user1@ibm.com" "user2@ibm.com" --start "2026-09-28" --end "2026-09-30" --compare --common-only --duration 60

# Load list of attendees from a file
py -3 scripts/outlook_skill.py check-availability --file "attendees.txt" --date "2026-09-28" --compare
```
- Aliases: `free-busy`, `freebusy`, `check-calendar`
- Queries Exchange / Outlook Free/Busy data across the organization for all specified attendees
- `--date`: Target date (YYYY-MM-DD)
- `--start` / `--end`: Date range (YYYY-MM-DD)
- `--days`: Number of days to check (default: 1)
- `--interval`: Slot resolution in minutes (`15`, `30`, `60`, default: `30`)
- `--working-hours`: Filter working window (default: `09:00-18:00`, or `all`)
- `--duration` / `--min-duration`: Minimum required meeting duration in minutes (e.g. `30`, `45`, `60`)
- `--common-only` / `--slots-only`: Show only the mutual common free slots list (omits individual calendar breakdowns)
- `--file`: Path to a text or CSV file with email addresses (comma-separated or one per line)
- `--compare`: Cross-references with your own calendar to calculate **Mutual Common Free Slots**
- `--json`: Machine-readable structured JSON output

### Respond to Meeting
```bash
py -3 scripts/outlook_skill.py respond-meeting "<meeting_entry_id>" --action accept
py -3 scripts/outlook_skill.py respond-meeting "<meeting_entry_id>" --action tentative
py -3 scripts/outlook_skill.py respond-meeting "<meeting_entry_id>" --action decline
py -3 scripts/outlook_skill.py respond-meeting "<meeting_entry_id>" --action propose --start "2026-06-25 14:00" --end "2026-06-25 15:00"
```
- Responds to a meeting invite in your inbox
- `--action`: accept, tentative, decline, or propose
- `--start`/`--end`: Required for `propose` (format: `YYYY-MM-DD HH:MM`)
- Sends response immediately (no draft gate — user explicitly chose the action)
- Prints confirmation: action, subject, organizer, time

### Sending Inline Images
```bash
py -3 assistant_brain/skills/outlook_com_skill/scripts/outlook_skill.py compose --to "email" --subject "text" --body-file "outbox/<name>.html" --inline-image "C:/path/image.png:myimage"
```
- `--inline-image`: Format is `filepath:cid_name` (comma separated for multiple)
- Reference in HTML body via `<img src="cid:cid_name">`
- Works with `compose`, `reply`, `forward`, `redirect`

### Viewing Embedded Images
When search results show `🖼 Embedded images (N): filename.png, ...`:
1. Use `get-email "<id>"` — images auto-save to temp with paths printed
2. Use Read tool on the printed path to view the image
3. AI describes or acts on image content

No extra user command needed — AI handles download + viewing transparently.

## Configuration

All configuration is centralized in [`backend/config.py`](backend/config.py).

**To change batch size:**
Edit `backend/config.py` and modify:
```python
class BatchConfig:
    OUTLOOK_BCC_LIMIT = 500  # Change this value
```

**Batch size recommendations:**
- **500** (default): Recommended for production use
- **100**: For testing with smaller batches
- **1000**: Maximum (may hit Exchange server limits)

## HTML Format Examples

```html
Hi John,<br><br>
Message text here.<br>
Second line of the message.
```

**Lists — NO `<ul><li>`:** Outlook adds excessive spacing between list items. Use `<br>` with plain ASCII hyphens (`-`) instead of decorative bullets:

```html
<!-- ✅ CORRECT -->
<p>- Item one<br>
- Item two<br>
- Item three</p>

<!-- ❌ WRONG: huge gaps in Outlook -->
<ul><li>Item one</li><li>Item two</li></ul>
```

**Style rule:** Prefer plain ASCII hyphens (`-`) and straight apostrophes (`'`) in outbound email bodies unless special characters are necessary. Body transport rules: see [Body Transport](#body-transport-all-send-commands).

## Find Workflow for Email Addresses

1. Lookup display name: `lookup-contact "user@domain.com"`
2. Find by display name: `find --from "Display Name"`

**Why:** Outlook MAPI doesn't reliably search by email address

## Recommended AI Usage Flow

### Finding All Emails About a Topic
```bash
# Step 1: Start narrow and recent with a specific query
py -3 scripts/outlook_skill.py find --subject "voucher approval" --folders "Inbox,Sent Items" --days 14

# Step 2: If not found, widen the time window but make the query more specific
py -3 scripts/outlook_skill.py find --subject "voucher approval philippines" --folders "Inbox,Sent Items" --days 45

# Step 3: From any result, find the full thread
py -3 scripts/outlook_skill.py find-thread "<entry_id>"

# Step 4: For even more context, find related across threads
py -3 scripts/outlook_skill.py find-related "<entry_id>"
```

## Custom Folders

| Folder | Path | Purpose |
|--------|------|---------|
| Recognized | `luomn@cn.ibm.com/Inbox/Recognized` | Thanks@IBM recognition emails |

**Note:** `move-email` cannot resolve subfolder names alone (e.g., `"Recognized"` fails). Always use the full mailbox path: `luomn@cn.ibm.com/Inbox/FolderName`.

## Requirements

- Microsoft Outlook 2016+ (running)
- Windows 10+
- Python 3.8+ with pywin32
- SQLite 3.35+ (included with Python 3.8+)
