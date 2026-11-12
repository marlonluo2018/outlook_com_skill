"""Shared helpers for Outlook CLI commands."""

from __future__ import annotations

import base64
import html
import json
import os
import re
import sys
import tempfile
import time

from backend.config import HostPathConfig


IMAGE_EXTENSIONS = {'.png', '.jpg', '.jpeg', '.gif', '.bmp', '.svg', '.ico', '.webp'}
DOCUMENT_EXTENSIONS = {'.pdf', '.doc', '.docx', '.xls', '.xlsx', '.ppt', '.pptx', '.txt', '.zip', '.rar'}
_FONT_STYLE = 'font-family:Calibri,sans-serif;font-size:11pt;'


_SEARCH_CACHE_NAME = 'last-find.json'
_SEARCH_CACHE_STALE_AFTER_MINUTES = 60


def _repo_root():
    """Best-effort repo root (the CLI lives under assistant_brain/skills/...)."""
    here = os.path.abspath(os.path.dirname(__file__))
    for _ in range(6):
        if os.path.isdir(os.path.join(here, 'assistant_brain')):
            return here
        here = os.path.dirname(here)
    return os.getcwd()


def _search_cache_path():
    cache_dir = str(HostPathConfig.cache_dir())
    try:
        os.makedirs(cache_dir, exist_ok=True)
    except Exception:
        cache_dir = _repo_root()
    return os.path.join(cache_dir, _SEARCH_CACHE_NAME)


def cache_search_results(emails, context=''):
    """Remember the last displayed result set so callers can reference row #N
    instead of copying an EntryID.

    `context` records which command produced the rows (e.g. "find --from 'x'
    --days 7"). It is echoed back whenever the cache is used, so referencing
    a stale or different-than-expected result set is visible instead of silent.
    """
    rows = []
    for idx, e in enumerate(emails or [], 1):
        rows.append({
            'n': idx,
            'entry_id': e.get('id') or e.get('entry_id', ''),
            'subject': e.get('subject', ''),
            'sender': e.get('sender', ''),
            'received': str(e.get('received_time', ''))[:19],
            'folder': e.get('folder', ''),
        })
    payload = {
        'rows': rows,
        'context': context or 'find',
        'saved_at': time.time(),
    }
    try:
        with open(_search_cache_path(), 'w', encoding='utf-8') as fh:
            json.dump(payload, fh, ensure_ascii=False, indent=1)
    except Exception:
        pass


def load_search_cache():
    try:
        with open(_search_cache_path(), 'r', encoding='utf-8') as fh:
            return json.load(fh).get('rows', [])
    except Exception:
        return []


def load_search_cache_provenance():
    """Source command and save time of the cached result set."""
    try:
        with open(_search_cache_path(), 'r', encoding='utf-8') as fh:
            data = json.load(fh)
        return {
            'context': data.get('context', 'unknown find'),
            'saved_at': data.get('saved_at', 0),
        }
    except Exception:
        return {'context': 'unknown find', 'saved_at': 0}


def cache_provenance_note():
    """One-line provenance for the cached result set, shown on --pick/--last.

    The cache holds only the most recent result set, so a caller can think it
    references one `find` while the cache actually holds another. Printing the
    source command and its age makes that mismatch visible instead of silent;
    warn once the cache is older than _SEARCH_CACHE_STALE_AFTER_MINUTES.
    """
    meta = load_search_cache_provenance()
    note = f"from `{meta['context']}`"
    if not meta['saved_at']:
        return note + " (age unknown)"
    age_min = max(0, (time.time() - meta['saved_at']) / 60)
    if age_min >= _SEARCH_CACHE_STALE_AFTER_MINUTES:
        return note + f" ⚠ {int(age_min)} min old — if this is not the result set you meant, re-run find"
    return note + f" ({int(age_min)} min ago)"


MIN_ENTRY_ID_LENGTH = 140


def _require_full_entry_id(email_id, subject=None):
    """Reject a truncated or hallucinated EntryID before it reaches MAPI.

    A valid Outlook mail EntryID is ~140 hex characters. A shorter one was
    truncated (terminal-wrapped hand copy) or hallucinated (an LLM asked to
    reproduce the ID drops characters — observed 160 -> 90 with DeepSeek V4)
    and can never be opened, so fail fast with an actionable message instead
    of a cryptic COM error. Skipped when a --subject fallback is available.
    """
    id_text = str(email_id or '').strip()
    if len(id_text) < MIN_ENTRY_ID_LENGTH and not subject:
        raise ValueError(
            f"EntryID rejected: {len(id_text)} characters, but a valid Outlook mail "
            f"EntryID is about {MIN_ENTRY_ID_LENGTH}. A short ID is a truncated or "
            "hallucinated copy that can never be opened. Reference the email instead: "
            "get-email --pick N / --last (row from the last `find`), get-email --task "
            "T### [--pick N] (task-file marker), or pass --subject."
        )


_FALLBACK_FOLDER_IDS = (6, 5, 16, 3, 4)  # Inbox, Sent, Drafts, Deleted, Outbox
_FALLBACK_SCAN_LIMIT = 1500


def _scan_folder_for_entry_id(folder, email_id, limit):
    """Return the item whose EntryID matches, scanning the folder newest-first."""
    try:
        items = folder.Items
    except Exception:
        return None
    try:
        items.Sort("[ReceivedTime]", True)
    except Exception:
        pass
    try:
        item = items.GetFirst()
    except Exception:
        return None
    seen = 0
    while item is not None and seen < limit:
        seen += 1
        try:
            if item.EntryID == email_id:
                return item
        except Exception:
            pass
        try:
            item = items.GetNext()
        except Exception:
            return None
    return None


def _resolve_by_enumeration(session, email_id):
    """Resolve an item by scanning well-known folders.

    Some Outlook/MAPI sessions reject GetItemFromID for every EntryID even
    though folder enumeration reads the same items fine. Enumeration costs a
    couple of seconds for a few hundred items, so it is a safe fallback.
    """
    try:
        namespace = session.namespace or session.outlook.GetNamespace("MAPI")
    except Exception:
        return None
    for folder_id in _FALLBACK_FOLDER_IDS:
        try:
            folder = namespace.GetDefaultFolder(folder_id)
        except Exception:
            continue
        found = _scan_folder_for_entry_id(folder, email_id, _FALLBACK_SCAN_LIMIT)
        if found is not None:
            print(f"[outlook] GetItemFromID failed; resolved by scanning {folder.Name}")
            return found
    return None


_WARM_FOLDER_IDS = (6, 5)  # Inbox, Sent Items
_WARM_SAMPLE = 25


def _warm_store(namespace):
    """Touch a few recent items so MAPI materialises the store table.

    In this environment GetItemFromID can fail on a cold session even for a
    perfectly valid EntryID, and succeed once the store has been read.
    """
    for folder_id in _WARM_FOLDER_IDS:
        try:
            items = namespace.GetDefaultFolder(folder_id).Items
            items.Sort("[ReceivedTime]", True)
            item = items.GetFirst()
            for _ in range(_WARM_SAMPLE):
                if item is None:
                    break
                item = items.GetNext()
        except Exception:
            continue


def _resolve_by_subject(session, subject):
    """Find the newest item whose Subject matches exactly.

    Subject filters (Items.Restrict) work reliably in sessions where
    GetItemFromID does not, so this is the dependable escape hatch: pass the
    subject shown by find/find-recent when the EntryID cannot be opened.
    """
    try:
        namespace = session.namespace or session.outlook.GetNamespace("MAPI")
    except Exception:
        return None
    quoted = str(subject).replace('"', '""')
    filter_expr = '[Subject] = "' + quoted + '"'
    for folder_id in _FALLBACK_FOLDER_IDS:
        try:
            folder = namespace.GetDefaultFolder(folder_id)
        except Exception:
            continue
        try:
            items = folder.Items
            try:
                items.Sort("[ReceivedTime]", True)
            except Exception:
                pass
            matches = items.Restrict(filter_expr)
            if matches.Count:
                print(f"[outlook] resolved by subject in {folder.Name}: {subject[:60]}")
                return matches.Item(1)
        except Exception:
            continue
    return None


def _get_email_item(session, email_id, subject=None):
    """Get email by ID (optionally by subject when the ID cannot be opened)."""
    _require_full_entry_id(email_id, subject=subject)
    namespace = session.namespace or session.outlook.GetNamespace("MAPI")

    try:
        return namespace.GetItemFromID(email_id)
    except Exception as first_error:
        last_error = first_error

        try:
            stores = getattr(namespace, "Folders", None)
            if stores:
                for i in range(1, stores.Count + 1):
                    try:
                        store_root = stores.Item(i)
                        store_id = getattr(store_root, "StoreID", "")
                        if not store_id:
                            continue
                        return namespace.GetItemFromID(email_id, store_id)
                    except Exception as store_error:
                        last_error = store_error
                        continue
        except Exception:
            pass

        # Cold-session MAPI quirk: read the store first, then retry the ID.
        _warm_store(namespace)
        try:
            return namespace.GetItemFromID(email_id)
        except Exception as retry_error:
            last_error = retry_error

        # Subject lookup is the dependable path in this environment.
        if subject:
            found = _resolve_by_subject(session, subject)
            if found is not None:
                return found

        # Last resort: walk the well-known folders looking for the EntryID.
        found = _resolve_by_enumeration(session, email_id)
        if found is not None:
            return found

        if "moved or deleted" in str(last_error).lower():
            print("Error: Outlook could not resolve the email by its current item handle.")
            print("This can happen after recall, move, or mailbox-store changes.")
            print("Please search for the email again to get a fresh current email ID.")
        else:
            length = len(str(email_id).strip())
            print("Error: Outlook could not open that email ID.")
            print("ID length received: %d characters." % length)
            if length < 140:
                print("A valid Outlook EntryID is about 140 characters. Anything shorter is")
                print("truncated or short-term (session-scoped) and can never be opened from")
                print("a new process. Re-copy the full ID from `find`, or use --subject.")
            print("Tried GetItemFromID (default store + every store), a warmed retry, a")
            print("subject lookup (only when --subject is given) and a folder scan of")
            print("Inbox / Sent Items / Drafts / Deleted Items / Outbox.")

        raise last_error

def _add_attachments(mail_item, attach_str):
    """Add file attachments to a mail item."""
    if not attach_str:
        return
    import os
    for filepath in attach_str.split(","):
        filepath = filepath.strip().strip('"')
        if not os.path.exists(filepath):
            print(f"WARNING: Attachment not found: {filepath}")
            continue
        mail_item.Attachments.Add(filepath)

def _attach_emails(mail_item, attach_email_str, session):
    """Attach other emails as .msg files to a mail item."""
    if not attach_email_str:
        return
    import os, tempfile
    temp_files = []
    for eid in attach_email_str.split(","):
        eid = eid.strip().strip('"')
        if not eid:
            continue
        try:
            source = _get_email_item(session, eid)
            subject = str(getattr(source, "Subject", "email"))[:50]
            safe_name = "".join(c if c.isalnum() or c in " -_" else "_" for c in subject)
            temp_path = os.path.join(tempfile.gettempdir(), f"{safe_name}.msg")
            source.SaveAs(temp_path, 3)  # olMSG
            mail_item.Attachments.Add(temp_path)
            temp_files.append(temp_path)
        except Exception as e:
            print(f"WARNING: Could not attach email {eid[:20]}...: {e}")
    for f in temp_files:
        try:
            os.remove(f)
        except OSError:
            pass

def _set_importance(mail_item, importance_str):
    """Set email importance/priority flag. Values: high, low, normal (default)."""
    if not importance_str:
        return
    level = importance_str.strip().lower()
    if level == "high":
        mail_item.Importance = 2  # olImportanceHigh
    elif level == "low":
        mail_item.Importance = 0  # olImportanceLow

def _add_inline_images(mail_item, inline_str):
    """Embed images into the HTML body using CID.

    Format: "filepath:cid_name" (comma separated for multiple).
    Windows paths (C:\\...) are handled -- only the last colon past position 1
    is treated as the filepath:cid separator.

    If the body already contains cid: references for the images, only sets the
    CID property (user controls placement). Otherwise auto-prepends <img> tags.
    """
    if not inline_str:
        return
    import os
    cids = []
    for entry in inline_str.split(","):
        entry = entry.strip().strip('"')
        last_colon = entry.rfind(":")
        if last_colon > 1:
            filepath = entry[:last_colon]
            cid = entry[last_colon + 1:]
        else:
            filepath = entry
            cid = None
        if not os.path.exists(filepath):
            print(f"WARNING: Inline image not found: {filepath}")
            continue
        if not cid:
            cid = os.path.basename(filepath).replace(" ", "_")
        # Remove existing attachments with the same CID (e.g. from forwarded emails)
        for i in range(mail_item.Attachments.Count, 0, -1):
            try:
                existing_cid = mail_item.Attachments.Item(i).PropertyAccessor.GetProperty(
                    "http://schemas.microsoft.com/mapi/proptag/0x3712001F"
                )
                if existing_cid == cid:
                    mail_item.Attachments.Remove(i)
            except Exception:
                pass
        attachment = mail_item.Attachments.Add(filepath)
        attachment.PropertyAccessor.SetProperty(
            "http://schemas.microsoft.com/mapi/proptag/0x3712001F", cid
        )
        cids.append(cid)
    if cids and not any(f"cid:{c}" in mail_item.HTMLBody for c in cids):
        img_html = "".join(f'<img src="cid:{c}" style="max-width:100%;"><br>' for c in cids)
        mail_item.HTMLBody = img_html + mail_item.HTMLBody
    elif cids:
        mail_item.HTMLBody = mail_item.HTMLBody

def _preserve_inline_images(dest_item, source_item):
    """Copy CID-referenced inline images from source to dest to preserve them in replies.

    After HTMLBody reassignment, Outlook drops inline attachments whose CID
    references it can no longer correlate. This re-adds them explicitly.
    """
    import tempfile
    PR_ATTACH_CONTENT_ID = "http://schemas.microsoft.com/mapi/proptag/0x3712001F"

    # Collect CIDs already on the dest item to avoid duplicates
    existing_cids = set()
    for i in range(1, dest_item.Attachments.Count + 1):
        try:
            cid = dest_item.Attachments.Item(i).PropertyAccessor.GetProperty(PR_ATTACH_CONTENT_ID)
            if cid:
                existing_cids.add(cid)
        except Exception:
            pass

    added = False
    for i in range(1, source_item.Attachments.Count + 1):
        attach = source_item.Attachments.Item(i)
        try:
            cid = attach.PropertyAccessor.GetProperty(PR_ATTACH_CONTENT_ID)
        except Exception:
            cid = None
        if not cid:
            continue
        # Only copy if the CID is referenced in the HTML body and not already present
        if f"cid:{cid}" not in dest_item.HTMLBody:
            continue
        if cid in existing_cids:
            continue
        temp_path = os.path.join(tempfile.gettempdir(), attach.FileName or f"inline_{i}.png")
        try:
            attach.SaveAsFile(temp_path)
            new_attach = dest_item.Attachments.Add(temp_path)
            new_attach.PropertyAccessor.SetProperty(PR_ATTACH_CONTENT_ID, cid)
            added = True
        except Exception:
            pass
        finally:
            if os.path.exists(temp_path):
                os.remove(temp_path)

    if added:
        # Force Outlook to re-resolve CID references
        dest_item.HTMLBody = dest_item.HTMLBody

def _is_embedded_image(attachment):
    """Detect if an attachment is an embedded/inline image (4-method detection)."""
    file_name = getattr(attachment, 'FileName', '') or getattr(attachment, 'DisplayName', 'Unknown')
    ext = os.path.splitext(file_name)[1].lower()

    if ext in DOCUMENT_EXTENSIONS:
        return False
    if ext not in IMAGE_EXTENSIONS:
        return False

    # Method 1: Content-ID / Content-Location
    try:
        if hasattr(attachment, 'PropertyAccessor'):
            content_id = attachment.PropertyAccessor.GetProperty(
                "http://schemas.microsoft.com/mapi/proptag/0x3712001F")
            if content_id and str(content_id).strip():
                return True
            content_loc = attachment.PropertyAccessor.GetProperty(
                "http://schemas.microsoft.com/mapi/proptag/0x3713001F")
            if content_loc and str(content_loc).strip():
                return True
    except Exception:
        pass

    # Method 2: Attachment type (3=Embedded, 4=OLE)
    att_type = getattr(attachment, 'Type', 1)
    if att_type in (3, 4):
        return True

    # Method 3: Filename heuristics
    lower_name = file_name.lower()
    if any(p in lower_name for p in ('image', 'img', 'cid:', 'embedded')):
        return True
    stem = os.path.splitext(lower_name)[0]
    if stem.isdigit() or (len(stem) <= 2 and stem.isalnum()):
        return True

    # Method 4: Small image size (< 10 KB)
    size = getattr(attachment, 'Size', 0)
    if 0 < size < 10240:
        return True

    return False

def _wrap_body_font(html_body: str) -> str:
    """Wrap body content in a styled div.

    Plain text input is HTML-escaped and newlines become <br> so emails
    preserve formatting when assigned to Outlook HTMLBody.
    """
    if not html_body:
        return ""

    # If caller passed plain text, preserve visible line breaks in HTML mail.
    if not re.search(r'<[a-zA-Z][^>]*>', html_body):
        html_body = html.escape(html_body)
        html_body = html_body.replace('\r\n', '\n')
        html_body = html_body.replace('\n', '<br>')

    return f'<div style="{_FONT_STYLE}">{html_body}</div>'

def _format_forward_message_html(message_text: str) -> str:
    """Convert plain text or HTML-ish input into the simple prepended block used for forwards.
    If input already contains HTML tags, pass through as-is (no <p> wrapping)."""
    if not message_text:
        return ""
    if re.search(r'<[a-zA-Z][^>]*>', message_text):
        return _wrap_body_font(message_text)
    inner = message_text.replace('\n\n', '<br><br>').replace('\n', '<br>')
    return _wrap_body_font(inner)

def _ensure_utf8_charset(html: str) -> str:
    """Ensure HTML declares UTF-8 charset. Replace any existing charset or add one."""
    charset_pattern = re.compile(
        r'<meta[^>]*charset=["\']?[^"\'>]+["\']?[^>]*/?>',
        re.IGNORECASE
    )
    content_type_pattern = re.compile(
        r'<meta[^>]*http-equiv=["\']?Content-Type["\']?[^>]*/?>',
        re.IGNORECASE
    )
    utf8_meta = '<meta http-equiv="Content-Type" content="text/html; charset=utf-8">'

    if charset_pattern.search(html):
        html = charset_pattern.sub(utf8_meta, html, count=1)
    elif content_type_pattern.search(html):
        html = content_type_pattern.sub(utf8_meta, html, count=1)
    else:
        head_match = re.search(r'(<head[^>]*>)', html, re.IGNORECASE)
        if head_match:
            insert_pos = head_match.end()
            html = html[:insert_pos] + utf8_meta + html[insert_pos:]
    return html

def _inject_before_body(prepend_html: str, full_html: str) -> str:
    """Insert prepend_html inside the <body> tag and ensure UTF-8 charset."""
    full_html = _ensure_utf8_charset(full_html)
    body_match = re.search(r'(<body[^>]*>)', full_html, re.IGNORECASE)
    if body_match:
        insert_pos = body_match.end()
        return full_html[:insert_pos] + prepend_html + full_html[insert_pos:]
    return prepend_html + full_html

def _get_smtp_address(recipient):
    """Extract SMTP email address from an Outlook Recipient object."""
    try:
        if recipient.AddressEntry.Type == "EX":
            exchange_user = recipient.AddressEntry.GetExchangeUser()
            if exchange_user:
                return exchange_user.PrimarySmtpAddress
    except:
        pass
    return recipient.Address or recipient.Name

def _remove_self_from_recipients(reply, current_user_email):
    """Remove current user from reply recipients."""
    to_remove = []
    for i in range(1, reply.Recipients.Count + 1):
        recipient = reply.Recipients.Item(i)
        recipient_email = _get_smtp_address(recipient)
        if recipient_email and recipient_email.lower() == current_user_email.lower():
            to_remove.append(i)
    for i in reversed(to_remove):
        reply.Recipients.Remove(i)

def _copy_sent_items_recipients(email_item, reply, current_user_email, to_only=False):
    """Copy recipients from a Sent Items email onto a newly created reply item.

    Excludes the current user (who was the sender of the original email).
    With to_only=True, copies only To (Type 1) recipients — used by --only
    replies, where the original To recipient is the thread counterpart.
    """
    for i in range(1, email_item.Recipients.Count + 1):
        recip = email_item.Recipients.Item(i)
        if to_only and getattr(recip, "Type", 1) != 1:  # 1 = olTo
            continue
        smtp = _get_smtp_address(recip)
        if smtp and smtp.lower() != current_user_email.lower():
            new_recip = reply.Recipients.Add(smtp)
            new_recip.Type = getattr(recip, "Type", 1) or 1  # 1=To, 2=CC

def _normalize_cc(cc_arg):
    """Normalize args.cc (list from append or single string) into one comma-separated string."""
    if not cc_arg:
        return ""
    if isinstance(cc_arg, list):
        return ",".join(cc_arg)
    return cc_arg

def _add_recipients(reply, to_str, cc_arg):
    """Append --to and --cc to existing recipients."""
    if to_str:
        for r in to_str.replace(";", ",").split(","):
            r = r.strip()
            if r:
                reply.Recipients.Add(r)
    cc_str = _normalize_cc(cc_arg)
    if cc_str:
        for r in cc_str.replace(";", ",").split(","):
            r = r.strip()
            if r:
                cc_recip = reply.Recipients.Add(r)
                cc_recip.Type = 2

def _build_reply_header(email_item, current_user_email):
    """Build standard reply attribution header (From/Sent/To/Cc/Subject)."""
    sent_on = email_item.SentOn.strftime("%A, %B %d, %Y %I:%M %p") if email_item.SentOn else ""
    original_to = getattr(email_item, "To", "") or ""
    original_cc = getattr(email_item, "CC", "") or ""
    original_subject = getattr(email_item, "Subject", "") or ""
    header = (
        '<div style="border:none;border-top:solid #E1E1E1 1.0pt;padding:3.0pt 0 0 0">'
        f'<p><b>From:</b> {email_item.SenderName} &lt;{current_user_email}&gt;<br>'
        f'<b>Sent:</b> {sent_on}<br>'
        f'<b>To:</b> {original_to}<br>'
    )
    if original_cc:
        header += f'<b>Cc:</b> {original_cc}<br>'
    header += f'<b>Subject:</b> {original_subject}</p></div>'
    return header

def _bare_subject(subject):
    return re.sub(r'^(RE:\s*|FW:\s*)+', '', subject or '', flags=re.IGNORECASE)

def _sent_on_to_utc(sent_on):
    from datetime import timezone
    if hasattr(sent_on, 'astimezone'):
        return sent_on.astimezone(timezone.utc)
    return sent_on.replace(tzinfo=timezone.utc)

def _collect_recent_sent_entry_ids(session, subject=None, seconds=300, max_items=100):
    """Snapshot recent Sent Items EntryIDs before sending.

    This prevents a post-send lookup from returning an older email with the
    same subject when repeated tests or retries happen within the lookup
    window.
    """
    from datetime import datetime, timedelta, timezone
    threshold = datetime.now(timezone.utc) - timedelta(seconds=seconds)
    bare_subject = _bare_subject(subject)
    entry_ids = set()
    try:
        ns = session.outlook.GetNamespace("MAPI")
        sent_folder = ns.GetDefaultFolder(5)  # 5 = olFolderSentMail
        items = sent_folder.Items
        items.Sort("[SentOn]", True)
        item = items.GetFirst()
        scanned = 0
        while item and scanned < max_items:
            scanned += 1
            sent_utc = _sent_on_to_utc(item.SentOn)
            if sent_utc < threshold:
                break
            if not subject or bare_subject in _bare_subject(getattr(item, 'Subject', '')):
                entry_id = getattr(item, 'EntryID', None)
                if entry_id:
                    entry_ids.add(entry_id)
            item = items.GetNext()
    except Exception:
        pass
    return entry_ids

def _print_sent_entry_id(session, subject, sent_after=None, exclude_entry_ids=None):
    """Retrieve and print EntryID of the newly sent email.

    If exclude_entry_ids is provided, only a Sent Items entry that was not
    present before Send() is eligible. This avoids printing the same EntryID
    for consecutive sends with the same subject.
    """
    import time
    from datetime import datetime, timedelta, timezone
    exclude_entry_ids = set(exclude_entry_ids or [])
    threshold = sent_after or (datetime.now(timezone.utc) - timedelta(seconds=90))
    # Outlook/Exchange timestamps can lag slightly; allow a tiny tolerance,
    # while exclude_entry_ids still blocks older already-seen messages.
    threshold = threshold - timedelta(seconds=5)
    bare_subject = _bare_subject(subject)
    for attempt in range(15):
        time.sleep(2)
        try:
            ns = session.outlook.GetNamespace("MAPI")
            sent_folder = ns.GetDefaultFolder(5)  # 5 = olFolderSentMail
            items = sent_folder.Items
            items.Sort("[SentOn]", True)
            item = items.GetFirst()
            while item:
                sent_on = item.SentOn
                sent_utc = _sent_on_to_utc(sent_on)
                if sent_utc < threshold:
                    break
                entry_id = getattr(item, 'EntryID', None)
                if entry_id in exclude_entry_ids:
                    item = items.GetNext()
                    continue
                bare_item_subject = _bare_subject(getattr(item, 'Subject', ''))
                if not subject or bare_subject in bare_item_subject:
                    local_time = sent_on if sent_on.tzinfo else sent_on.replace(tzinfo=None)
                    print(f"SentOn: {local_time.strftime('%a %b %d, %Y %H:%M')}")
                    print(f"EntryID: {entry_id}")
                    return entry_id
                item = items.GetNext()
        except Exception:
            pass
    print("Warning: sent EntryID lookup did not find a new matching Sent Items entry.", file=sys.stderr)
    return None

def _reply_step(message):
    print(f"[reply] {message}", file=sys.stderr, flush=True)

def _resolve_body(args, required=True):
    """Resolve body from transport mechanisms.

    Priority:
    1. --body-file
    2. --body-stdin-base64
    3. --body-base64
    4. --body-stdin
    5. auto-detect stdin redirection (base64 or UTF-8 plain text)
    6. direct body arg

    Returns True on success (args.body is set), False on error."""
    body_file = getattr(args, 'body_file', None)
    if body_file:
        try:
            from pathlib import Path
            p = Path(body_file)
            if p.is_file():
                args.body = p.read_text(encoding='utf-8')
                print(f"[resolve_body] loaded body from file: {body_file} ({len(args.body)} chars)", file=sys.stderr)
            else:
                print(f"Error: body file not found: {body_file}", file=sys.stderr)
                return False
        except Exception as exc:
            print(f"Error reading body file {body_file}: {exc}", file=sys.stderr)
            return False

    if getattr(args, 'body_stdin_base64', False):
        if not sys.stdin.isatty():
            try:
                raw = sys.stdin.buffer.read()
                args.body = base64.b64decode(raw).decode('utf-8')
            except Exception as exc:
                print(f"Error: invalid base64 stdin payload: {exc}", file=sys.stderr)
                return False
        else:
            print("Error: --body-stdin-base64 specified but no piped input detected.", file=sys.stderr)
            return False

    body_base64 = getattr(args, 'body_base64', None)
    if body_base64:
        try:
            args.body = base64.b64decode(body_base64).decode('utf-8')
        except Exception as exc:
            print(f"Error: invalid --body-base64 payload: {exc}", file=sys.stderr)
            return False

    if getattr(args, 'body_stdin', False):
        if not sys.stdin.isatty():
            try:
                args.body = sys.stdin.buffer.read().decode('utf-8')
            except UnicodeDecodeError as exc:
                print(f"Error: stdin is not valid UTF-8: {exc}", file=sys.stderr)
                return False
        else:
            print("Error: --body-stdin specified but no piped input detected.", file=sys.stderr)
            return False

    # Auto-detect stdin redirection if args.body is not set yet
    if not args.body and not sys.stdin.isatty():
        try:
            raw_bytes = sys.stdin.buffer.read()
            if raw_bytes:
                # Try to auto-detect if it's base64 encoded text
                try:
                    decoded_str = raw_bytes.decode('ascii', errors='strict')
                    # Base64 string characters check
                    cleaned_str = re.sub(r'\s+', '', decoded_str)
                    if re.match(r'^[A-Za-z0-9+/=]+$', cleaned_str) and len(cleaned_str) % 4 == 0:
                        try:
                            decoded_body = base64.b64decode(cleaned_str.encode('ascii')).decode('utf-8')
                            # Ensure it doesn't look like raw binary bytes (e.g. control chars)
                            if all(ord(c) >= 32 or c in '\r\n\t' for c in decoded_body):
                                args.body = decoded_body
                        except Exception:
                            pass
                except Exception:
                    pass

                # Fallback to direct UTF-8 decoding of the stream if base64 didn't resolve
                if not args.body:
                    try:
                        args.body = raw_bytes.decode('utf-8')
                    except UnicodeDecodeError as exc:
                        print(f"Error: redirected stdin is not valid UTF-8: {exc}", file=sys.stderr)
                        return False
        except Exception as exc:
            print(f"Error reading redirected stdin: {exc}", file=sys.stderr)
            return False

    if required and not args.body:
        print("Error: body or piped stdin is required.", file=sys.stderr)
        return False
    return True
