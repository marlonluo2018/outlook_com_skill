"""Formatting helpers for CLI email and thread output."""

from __future__ import annotations

import re

from backend.config import display_config
from backend.outlook_session.session_manager import OutlookSessionManager


def _folder_emoji(folder_name: str) -> str:
    """Return emoji indicator for folder type."""
    name_lower = folder_name.lower()
    if "sent" in name_lower:
        return "\U0001F4E4"  # 📤
    elif "inbox" in name_lower:
        return "\U0001F4E5"  # 📥
    elif "draft" in name_lower:
        return "\U0001F4DD"  # 📝
    elif "deleted" in name_lower or "trash" in name_lower:
        return "\U0001F5D1"  # 🗑
    return "\U0001F4C1"  # 📁

def _build_body_preview(body_text: str) -> str:
    """Build a compact terminal-safe preview from an email body."""
    if not body_text:
        return ""

    stop_markers = (
        'from:',
        'sent:',
        'subject:',
        'to:',
        'cc:',
        'original message',
        '-----original message-----',
        'zjqcmqryfpfptbannerstart',
        'zjqcmqryfpfptbannerend',
        'notice:this is an external sender',
        'this message is from an external sender',
    )

    preview_lines = []
    preview_budget = max(getattr(display_config, 'PREVIEW_LENGTH', 200) * 2, 200)

    for raw_line in body_text.splitlines():
        line = raw_line.strip()
        if not line:
            continue

        lower = line.lower()
        if any(marker in lower for marker in stop_markers):
            break
        if re.fullmatch(r'[_=\-\s]{5,}', line):
            break
        if line.startswith('<http') or line.startswith('http://') or line.startswith('https://'):
            continue
        if 'proofpoint.com' in lower or line == 'Report Suspicious':
            continue

        line = re.sub(r'\s+', ' ', line)
        if not line:
            continue

        preview_lines.append(line)

        if len(' '.join(preview_lines)) >= preview_budget or len(preview_lines) >= 4:
            break

    preview = re.sub(r'\s+', ' ', ' '.join(preview_lines)).strip()
    max_preview_length = getattr(display_config, 'PREVIEW_LENGTH', 200)
    if len(preview) > max_preview_length:
        preview = preview[:max_preview_length].rstrip() + "..."
    return preview

def _display_email_list(emails, show_folder=True):
    """Display formatted email list with folder markers."""
    def extract_display_name(recipient_string):
        if not recipient_string:
            return ""
        if '</o=' in recipient_string.lower() or '/cn=' in recipient_string.lower():
            if '<' in recipient_string:
                return recipient_string.split('<')[0].strip()
        if '<' in recipient_string and '@' in recipient_string:
            return recipient_string.split('<')[0].strip()
        return recipient_string.strip()

    for idx, email_data in enumerate(emails, 1):
        email_id = email_data.get('id') or email_data.get('entry_id', '')
        print(f"{'='*80}")
        folder_name = email_data.get('folder', '')
        folder_indicator = f" {_folder_emoji(folder_name)} {folder_name}" if show_folder and folder_name else ""

        meeting = email_data.get('meeting_status', '')
        meeting_icon = ""
        if meeting == "meeting_request":
            meeting_icon = " 📅 Meeting Invite"
        elif meeting == "meeting_canceled":
            meeting_icon = " ❌ Canceled"
        elif meeting == "meeting":
            meeting_icon = " 📅 Meeting"
        elif not meeting:
            # Fast check: subject + sender heuristics (no body access)
            subj_lower = (email_data.get('subject') or '').lower()
            sender_lower = (email_data.get('sender') or '').lower()
            event_sender_kw = ('events', 'calendar', 'webinar', 'noreply')
            event_subj_kw = ('webinar', 'join us', 'register now', 'you are invited',
                             "you're invited", 'invitation', 'save the date',
                             'live event', 'virtual event')
            if any(kw in sender_lower for kw in event_sender_kw) or \
               any(kw in subj_lower for kw in event_subj_kw):
                meeting_icon = " 📅 Event"

        print(f"Email #{idx}{folder_indicator}{meeting_icon}")
        print(f"{'='*80}")

        subject = email_data.get('subject', 'No Subject')
        sender = email_data.get('sender', 'Unknown')
        sender_email = email_data.get('sender_email')
        received = email_data.get('received_time', 'Unknown')

        print(f"ID: {email_id}")
        print(f"Subject: {subject}")
        if sender_email:
            print(f"From: {sender} <{sender_email}>")
        else:
            print(f"From: {sender}")

        to_recipients = email_data.get('to_recipients', [])
        if to_recipients:
            to_list = []
            for recipient in to_recipients:
                name = recipient.get('name', '')
                address = recipient.get('address', '')
                display_name = extract_display_name(name) if name else extract_display_name(address)
                if display_name:
                    if address and "@" in address:
                        to_list.append(f"{display_name} <{address}>")
                    else:
                        to_list.append(display_name)
            if to_list:
                print(f"To: {'; '.join(to_list)}")

        cc_recipients = email_data.get('cc_recipients', [])
        if cc_recipients:
            cc_list = []
            for recipient in cc_recipients:
                name = recipient.get('name', '')
                address = recipient.get('address', '')
                display_name = extract_display_name(name) if name else extract_display_name(address)
                if display_name:
                    if address and "@" in address:
                        cc_list.append(f"{display_name} <{address}>")
                    else:
                        cc_list.append(display_name)
            if cc_list:
                print(f"CC: {'; '.join(cc_list)}")

        print(f"Received: {received}")

        # Show confidence/strategy for related search
        confidence = email_data.get('_confidence')
        strategy = email_data.get('_strategy')
        if confidence is not None and strategy:
            stars = "★" * int(confidence * 5) + "☆" * (5 - int(confidence * 5))
            print(f"Relevance: {stars} ({strategy})")

        has_attachments = email_data.get('has_attachments', False)
        attachments = email_data.get('attachments', [])
        attachments_count = email_data.get('attachments_count', 0)
        embedded_images_count = email_data.get('embedded_images_count', 0)

        if has_attachments and attachments:
            print(f"\n📎 Attachments ({attachments_count}):")
            for attachment in attachments:
                name = attachment.get('name', 'Unknown')
                size = attachment.get('size', 0)
                size_kb = size / 1024 if size > 0 else 0
                print(f"  - {name} ({size_kb:.1f} KB)")

        if embedded_images_count > 0:
            embedded_images = email_data.get('embedded_images', [])
            if embedded_images:
                names = ", ".join(img.get('name', 'unknown') for img in embedded_images)
                print(f"\U0001F5BC  Embedded images ({embedded_images_count}): {names}")
            else:
                print(f"\U0001F5BC  Embedded images: {embedded_images_count}")

        try:
            with OutlookSessionManager() as session:
                email_item = session.outlook.GetNamespace("MAPI").GetItemFromID(email_id)
                body_text = email_item.Body if hasattr(email_item, 'Body') else ""
                preview = _build_body_preview(body_text)
                if preview:
                    print(f"\nPreview: {preview}")
        except Exception:
            pass

        print()

def _print_thread_summary(emails):
    """Print a one-line thread summary: count, participants, date span."""
    if not emails:
        return
    senders = list(dict.fromkeys(e.get('sender', 'Unknown') for e in emails))
    times = [e.get('received_time', '') for e in emails if e.get('received_time')]
    times.sort()

    participant_str = ", ".join(senders[:3])
    if len(senders) > 3:
        participant_str += f" +{len(senders) - 3} more"

    span_str = ""
    if len(times) >= 2:
        start = times[0][:10]
        end = times[-1][:10]
        if start != end:
            try:
                from datetime import datetime
                d1 = datetime.strptime(start, "%Y-%m-%d")
                d2 = datetime.strptime(end, "%Y-%m-%d")
                span_days = (d2 - d1).days
                span_str = f" | Span: {start} → {end} ({span_days} days)"
            except Exception:
                span_str = f" | Span: {start} → {end}"
        else:
            span_str = f" | Date: {start}"
    elif times:
        span_str = f" | Date: {times[0][:10]}"

    print(f"\U0001F4CA Thread: {len(emails)} messages | Participants: {participant_str}{span_str}")
    print()

def _format_recipients_brief(email_data):
    """Format To/CC as compact string: '→ Name1, Name2 +N'."""
    to_list = email_data.get('to_recipients', [])
    cc_list = email_data.get('cc_recipients', [])
    names = []
    for r in to_list:
        name = r.get('name', '') or r.get('address', '')
        if name:
            names.append(name.split('@')[0].split(',')[0].strip()[:15])
    total = len(to_list) + len(cc_list)
    if not names:
        return ""
    shown = names[:2]
    extra = total - len(shown)
    result = "→ " + ", ".join(shown)
    if extra > 0:
        result += f" +{extra}"
    return result

def _display_email_list_brief(emails, show_folder=True):
    """Compact single-line display with email ID on second line."""
    for idx, email_data in enumerate(emails, 1):
        email_id = email_data.get('id') or email_data.get('entry_id', '')
        folder_name = email_data.get('folder', '')
        folder_tag = f" {_folder_emoji(folder_name)}" if show_folder and folder_name else ""

        received = (email_data.get('received_time', '') or '')[:10]
        sender = (email_data.get('sender', 'Unknown') or 'Unknown')[:20]
        subject = (email_data.get('subject', 'No Subject') or 'No Subject')[:50]

        confidence = email_data.get('_confidence')
        strategy = email_data.get('_strategy', '')
        stars_str = ""
        if confidence is not None:
            stars = "★" * int(confidence * 5) + "☆" * (5 - int(confidence * 5))
            stars_str = f" {stars}"

        recipients_str = _format_recipients_brief(email_data)
        print(f"# {idx}{folder_tag} {received} | {sender:<20} | {subject}{stars_str}")
        print(f"  {recipients_str}  ID: {email_id}")
