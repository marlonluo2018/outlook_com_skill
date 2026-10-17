"""Read-only search, lookup, and email inspection commands."""

from __future__ import annotations

import json
import os
import re
import sys
import tempfile
from typing import Optional

from backend.config import path_config, search_config
from backend.email_search import (
    find_related_emails,
    find_thread_by_email_id,
    list_folders,
    list_recent_emails,
    list_recent_emails_multi,
    unified_search,
)
from backend.outlook_session.contact_operations import get_contact_by_email, get_contact_by_name
from backend.outlook_session.session_manager import OutlookSessionManager

from .common import _get_email_item, _get_smtp_address, _is_embedded_image
from .display import (
    _display_email_list,
    _display_email_list_brief,
    _print_thread_summary,
)


def _prepare_inline_image_dir(email_id: str) -> str:
    """Return the per-email inline-image folder under the project downloads
    dir, pruning folders whose files have all expired by age.

    Images are durable evidence (approval screenshots, report charts), so
    they live in the gitignored project downloads/outlook_inline/<EntryID>/
    instead of the OS temp dir; re-reading the same email reuses its folder.
    """
    import time

    base = path_config.inline_images_dir
    try:
        if base.is_dir():
            cutoff = time.time() - path_config.inline_images_max_age_days * 86400
            for folder in base.iterdir():
                if not folder.is_dir():
                    continue
                try:
                    files = [p for p in folder.iterdir() if p.is_file()]
                    if files and all(p.stat().st_mtime < cutoff for p in files):
                        for p in files:
                            p.unlink(missing_ok=True)
                        folder.rmdir()
                except OSError:
                    pass
        target = base / email_id[:16]
        target.mkdir(parents=True, exist_ok=True)
        return str(target)
    except OSError:
        # Unwritable project dir: fall back to a per-email OS temp folder
        import tempfile

        fallback = os.path.join(tempfile.gettempdir(), "outlook_inline", email_id[:16])
        os.makedirs(fallback, exist_ok=True)
        return fallback


def cmd_list_folders(args):
    """List all Outlook folders"""
    try:
        folders = list_folders(hide_system_folders=args.hide_system)
        print("\nAvailable folders:")
        for folder in folders:
            print(folder)
        return 0
    except Exception as e:
        print(f"Error: {str(e)}", file=sys.stderr)
        return 1

def _serialize_recent_email(email):
    """Return the stable JSON contract used by external sync consumers."""
    return {
        'entry_id': email.get('entry_id') or email.get('id', ''),
        'conversation_id': email.get('conversation_id', ''),
        'subject': email.get('subject', ''),
        'sender': email.get('sender', ''),
        'to_recipients': email.get('to_recipients', []),
        'cc_recipients': email.get('cc_recipients', []),
        'received_time': email.get('received_time', ''),
        'start_time': email.get('start_time', ''),
        'end_time': email.get('end_time', ''),
        'folder': email.get('folder_name', email.get('folder', '')),
        'has_attachments': email.get('has_attachments', False),
        'meeting_status': email.get('meeting_status', ''),
        'message_class': email.get('message_class', ''),
        'attachments_count': email.get('attachments_count', 0),
        'body_preview': email.get('body_preview', ''),
    }


def cmd_list_recent(args):
    """List recent emails with their IDs"""
    try:
        if args.folder:
            emails, message = list_recent_emails(args.folder, args.days, limit=args.limit, lightweight=getattr(args, 'json', False))
        else:
            emails, message = list_recent_emails_multi(days=args.days, limit=args.limit, lightweight=getattr(args, 'json', False))

        if args.limit and len(emails) > args.limit:
            emails = emails[:args.limit]

        email_count = len(emails)

        if getattr(args, 'json', False):
            serializable = [_serialize_recent_email(email) for email in emails]
            print(json.dumps(serializable, ensure_ascii=False))
            return 0

        print(f"\n✅ Found {email_count} recent emails\n")

        if emails:
            _display_email_list(emails, show_folder=True)

        return 0
    except Exception as e:
        print(f"Error: {str(e)}", file=sys.stderr)
        return 1

def _normalize_search_days(days: Optional[int]) -> int:
    """Normalize direct find search days while preserving broad-search capability."""
    if days is None:
        return search_config.DIRECT_FIND_DEFAULT_DAYS
    if days < 1:
        return 1
    return min(days, search_config.MAX_SEARCH_DAYS)

def cmd_search(args):
    """Search emails and display with IDs - supports multi-folder."""
    try:
        # Resolve field flags → (search_type, query) pairs
        field_searches = []
        if getattr(args, 'from_', None):
            field_searches.append(('sender', args.from_))
        if getattr(args, 'subject', None):
            field_searches.append(('subject', args.subject))
        if getattr(args, 'to', None):
            field_searches.append(('recipient', args.to))
        if getattr(args, 'body', None):
            field_searches.append(('body', args.body))

        # Legacy --type/--query fallback
        if not field_searches:
            if args.type and args.query:
                field_searches.append((args.type, args.query))
            else:
                print("Error: at least one search field required (--from, --subject, --to, or --body)", file=sys.stderr)
                return 1

        effective_days = _normalize_search_days(args.days)
        if args.days != effective_days:
            print(
                f"\nℹ️ Direct find search window adjusted to {effective_days} days "
                f"(allowed range: 1-{search_config.MAX_SEARCH_DAYS})."
            )

        # Resolve folders for multi-folder search
        if args.folders:
            folder_names = [f.strip() for f in args.folders.split(',')]
        else:
            folder_names = None

        # Single-field search (most common case)
        if len(field_searches) == 1:
            search_type, query = field_searches[0]
            # Auto-route '@' queries in subject/query to sender search
            if '@' in query and search_type == 'subject' and not getattr(args, 'subject', None):
                search_type = 'sender'

            emails, note = unified_search(
                search_term=query,
                days=effective_days,
                folder_name=args.folder,
                folder_names=folder_names,
                match_all=args.match_all,
                search_type=search_type,
            )

            # Auto-fallback for email address searches: if no results found in default window,
            # auto-expand search window to 90 days to find historical emails from that address.
            if not emails and '@' in query and search_type in ('sender', 'recipient') and args.days == search_config.DIRECT_FIND_DEFAULT_DAYS:
                expanded_days = 90
                emails, note = unified_search(
                    search_term=query,
                    days=expanded_days,
                    folder_name=args.folder,
                    folder_names=folder_names,
                    match_all=args.match_all,
                    search_type=search_type,
                )
                if emails:
                    note += f" (auto-expanded search window to {expanded_days} days for email address match)"
        else:
            # Multi-field: run each search, intersect by email ID
            result_sets = []
            for search_type, query in field_searches:
                emails_part, _ = unified_search(
                    search_term=query,
                    days=effective_days,
                    folder_name=args.folder,
                    folder_names=folder_names,
                    match_all=args.match_all,
                    search_type=search_type,
                )
                result_sets.append({e.get('id') or e.get('entry_id', ''): e for e in emails_part})

            # Intersect: keep only emails present in ALL result sets
            if result_sets:
                common_ids = set(result_sets[0].keys())
                for rs in result_sets[1:]:
                    common_ids &= set(rs.keys())
                emails = [result_sets[0][eid] for eid in common_ids if eid]
            else:
                emails = []
            fields_desc = " + ".join(f"{t}='{q}'" for t, q in field_searches)
            note = f"Found {len(emails)} emails matching {fields_desc}"

        if args.limit and len(emails) > args.limit:
            emails = emails[:args.limit]
            note += f" (showing first {args.limit})"

        print(f"\n✅ {note}\n")

        if emails:
            primary_type = field_searches[0][0]
            show_folder = (folder_names is not None and len(folder_names) > 1) or \
                          (folder_names is None and args.folder is None and primary_type == "subject")
            _display_email_list(emails, show_folder=show_folder)

        return 0
    except Exception as e:
        print(f"Error: {str(e)}", file=sys.stderr)
        return 1

def cmd_get_email(args):
    """Get full email details by ID"""
    try:
        email_ids = getattr(args, 'email_ids', [])
        if not email_ids and getattr(args, 'email_id_flag', None):
            email_ids = [args.email_id_flag]
        if not email_ids and getattr(args, 'email_id', None):
            if isinstance(args.email_id, list):
                email_ids = args.email_id
            else:
                email_ids = [args.email_id]

        if not email_ids:
            print("Error: email_id is required (positional or --id)", file=sys.stderr)
            return 1

        fields = None
        if getattr(args, 'fields', None):
            fields = set(f.strip().lower() for f in args.fields.split(','))

        truncate_len = getattr(args, 'truncate', None)
        latest_only = getattr(args, 'latest_only', False)
        max_lines = getattr(args, 'max_lines', None)

        with OutlookSessionManager() as session:
            for idx, email_id in enumerate(email_ids):
                if len(email_ids) > 1:
                    print("\n" + "="*40)
                    print(f"Email {idx+1}/{len(email_ids)}")
                    print("="*40)
                
                try:
                    email_item = _get_email_item(session, email_id)
                except Exception as e:
                    print(f"Error fetching email {email_id}: {str(e)}", file=sys.stderr)
                    continue

                # Detect item type from MessageClass
                message_class = getattr(email_item, 'MessageClass', 'IPM.Note') or 'IPM.Note'
                meeting_status = getattr(email_item, 'MeetingStatus', 0)
                is_meeting_item = message_class.startswith('IPM.Schedule.Meeting') or meeting_status > 0

                if not fields:
                    print("\nFull email details:")
                    print(f"ID: {email_id}")
                    # Show item type indicator
                    if message_class.startswith('IPM.Schedule.Meeting.Request'):
                        print("Type: New Invite")
                    elif message_class.startswith('IPM.Schedule.Meeting.Canceled'):
                        print("Type: Canceled")
                    elif message_class.startswith('IPM.Schedule.Meeting.Resp.Pos'):
                        print("Type: Accepted")
                    elif message_class.startswith('IPM.Schedule.Meeting.Resp.Neg'):
                        print("Type: Declined")
                    elif message_class.startswith('IPM.Schedule.Meeting.Resp.Tent'):
                        print("Type: Tentative")
                    elif message_class.startswith('IPM.Schedule.Meeting'):
                        print("Type: Meeting Update")
                    elif meeting_status > 0:
                        print("Type: Meeting")
                if not fields or 'subject' in fields:
                    print(f"Subject: {email_item.Subject}")
                if not fields or 'from' in fields:
                    try:
                        sender_name = email_item.SenderName or ''
                        sender_email = email_item.SenderEmailAddress or ''
                    except Exception:
                        sender_name = getattr(email_item, 'Organizer', '') or ''
                        sender_email = ''
                    # Resolve Exchange DN to SMTP address
                    if sender_email and (sender_email.startswith('/') or getattr(email_item, 'SenderEmailType', '') == 'EX'):
                        try:
                            smtp = email_item.Sender.GetExchangeUser().PrimarySmtpAddress
                            if smtp:
                                sender_email = smtp
                        except Exception:
                            sender_email = ''
                    if sender_email and sender_email != sender_name:
                        print(f"From: {sender_name} <{sender_email}>")
                    else:
                        print(f"From: {sender_name}")
                if not fields or 'to' in fields:
                    try:
                        to_recips = []
                        for i in range(1, email_item.Recipients.Count + 1):
                            recip = email_item.Recipients.Item(i)
                            if recip.Type == 1: # olTo
                                name = recip.Name or ''
                                email = _get_smtp_address(recip)
                                if email and email != name:
                                    to_recips.append(f"{name} <{email}>")
                                else:
                                    to_recips.append(name)
                        if to_recips:
                            print(f"To: {'; '.join(to_recips)}")
                        else:
                            to_val = email_item.To
                            if to_val:
                                print(f"To: {to_val}")
                    except Exception:
                        try:
                            to_val = email_item.To
                            if to_val:
                                print(f"To: {to_val}")
                        except Exception:
                            print("To: (unavailable — calendar item)")
                if not fields or 'cc' in fields:
                    try:
                        cc_recips = []
                        for i in range(1, email_item.Recipients.Count + 1):
                            recip = email_item.Recipients.Item(i)
                            if recip.Type == 2: # olCC
                                name = recip.Name or ''
                                email = _get_smtp_address(recip)
                                if email and email != name:
                                    cc_recips.append(f"{name} <{email}>")
                                else:
                                    cc_recips.append(name)
                        if cc_recips:
                            print(f"CC: {'; '.join(cc_recips)}")
                        else:
                            cc_val = email_item.CC
                            if cc_val:
                                print(f"CC: {cc_val}")
                    except Exception:
                        try:
                            cc_val = email_item.CC
                            if cc_val:
                                print(f"CC: {cc_val}")
                        except Exception:
                            pass
                if not fields or 'date' in fields:
                    rt = getattr(email_item, 'ReceivedTime', None)
                    print(f"Received: {rt.replace(tzinfo=None) if rt else 'Unknown'}")
                # Meeting-specific fields
                if is_meeting_item and (not fields or 'meeting' in fields):
                    _MAPI_START = 'http://schemas.microsoft.com/mapi/id/{00062002-0000-0000-C000-000000000046}/820D0040'
                    _MAPI_END = 'http://schemas.microsoft.com/mapi/id/{00062002-0000-0000-C000-000000000046}/820E0040'
                    # Try GetAssociatedAppointment for full meeting details
                    appt = None
                    if message_class.startswith('IPM.Appointment'):
                        appt = email_item
                    else:
                        try:
                            appt = email_item.GetAssociatedAppointment(False)
                        except Exception:
                            pass
                    if appt:
                        start_val = getattr(appt, 'Start', None)
                        end_val = getattr(appt, 'End', None)
                        if start_val:
                            s = start_val.replace(tzinfo=None)
                            e = end_val.replace(tzinfo=None) if end_val else None
                            time_line = str(s)
                            if e:
                                time_line += f" ~ {e}"
                            print(f"When: {time_line}")
                        loc = getattr(appt, 'Location', None)
                        if loc:
                            print(f"Location: {loc}")
                        organizer = getattr(appt, 'Organizer', None)
                        if organizer:
                            print(f"Organizer: {organizer}")
                        # Resolve attendees with email addresses via Recipients collection
                        try:
                            required = []
                            optional = []
                            for i in range(1, appt.Recipients.Count + 1):
                                recip = appt.Recipients.Item(i)
                                name = recip.Name or ''
                                email = _get_smtp_address(recip)
                                if email and email != name:
                                    entry = f"{name} <{email}>"
                                else:
                                    entry = name
                                if recip.Type == 1:  # olRequired
                                    required.append(entry)
                                elif recip.Type == 2:  # olOptional
                                    optional.append(entry)
                            if required:
                                print(f"Required: {'; '.join(required)}")
                            if optional:
                                print(f"Optional: {'; '.join(optional)}")
                        except Exception:
                            attendees = getattr(appt, 'RequiredAttendees', None)
                            if attendees:
                                print(f"Attendees: {attendees}")
                        resp_map = {0: "None", 1: "Organized", 2: "Tentative", 3: "Accepted", 4: "Declined", 5: "Not Responded"}
                        resp = getattr(appt, 'ResponseStatus', None)
                        if resp is not None and resp in resp_map:
                            print(f"Response: {resp_map[resp]}")
                    else:
                        # Fallback to MAPI properties on the inbox item
                        try:
                            pa = email_item.PropertyAccessor
                            start_val = pa.GetProperty(_MAPI_START)
                            end_val = pa.GetProperty(_MAPI_END)
                            if start_val:
                                s = start_val.replace(tzinfo=None)
                                e = end_val.replace(tzinfo=None) if end_val else None
                                time_line = str(s)
                                if e:
                                    time_line += f" ~ {e}"
                                print(f"When: {time_line}")
                        except Exception:
                            pass
                if not fields or 'body' in fields:
                    body = getattr(email_item, 'Body', '') or ''
                    if latest_only:
                        delimiters = [
                            r'\r?\n_{5,}',
                            r'\r?\n-----Original Message-----',
                            r'\r?\nFrom:\s+[^\r\n]+\r?\nSent:',
                            r'\r?\nFrom:\s+[^\r\n]+\r?\nDate:',
                        ]
                        pattern = '|'.join(delimiters)
                        parts = re.split(pattern, body, maxsplit=1, flags=re.IGNORECASE)
                        if parts:
                            body = parts[0].rstrip()
                    if max_lines is not None:
                        lines = body.splitlines()
                        if len(lines) > max_lines:
                            body = '\n'.join(lines[:max_lines]) + f"\n\n... [TRUNCATED to {max_lines} lines]"
                    if truncate_len is not None and len(body) > truncate_len:
                        truncated_body = body[:truncate_len] + f"\n\n... [TRUNCATED to {truncate_len} chars]"
                    else:
                        truncated_body = body
                    if body.strip():
                        print(f"\nBody:\n{truncated_body}")
                    elif is_meeting_item:
                        print("\nBody: (empty — meeting details are stored in the calendar item, check Calendar for full info)")
                    else:
                        print(f"\nBody:\n{truncated_body}")

                attachments_count = 0
                try:
                    attachments_count = email_item.Attachments.Count
                except Exception:
                    pass
                if (not fields or 'body' in fields) and attachments_count > 0:
                    regular = []
                    embedded = []
                    for i in range(1, email_item.Attachments.Count + 1):
                        attach = email_item.Attachments.Item(i)
                        if _is_embedded_image(attach):
                            embedded.append(attach)
                        else:
                            regular.append(attach)

                    if regular:
                        print(f"\n\u1F4CE Attachments ({len(regular)}):")
                        for attach in regular:
                            size_kb = attach.Size / 1024
                            print(f"  - {attach.FileName} ({size_kb:.1f} KB)")

                    if embedded:
                        inline_dir = _prepare_inline_image_dir(email_id)
                        print(f"\n\u1F5BC  Embedded images (auto-saved):")
                        used_names = set()
                        for attach in embedded:
                            filename = attach.FileName
                            if filename in used_names:
                                stem, ext = os.path.splitext(filename)
                                counter = 2
                                while f"{stem}_{counter}{ext}" in used_names:
                                    counter += 1
                                filename = f"{stem}_{counter}{ext}"
                            used_names.add(filename)
                            save_path = os.path.join(inline_dir, filename)
                            if not os.path.exists(save_path):
                                attach.SaveAsFile(save_path)
                            size_kb = attach.Size / 1024
                            print(f"  - {save_path} ({size_kb:.1f} KB)")

        return 0
    except Exception as e:
        print(f"Error: {str(e)}", file=sys.stderr)
        return 1

def cmd_lookup_contact(args):
    """Look up contact information by email address or display name"""
    try:
        query = args.query
        if '@' in query:
            contact_info = get_contact_by_email(query)
            results = [contact_info] if contact_info else []
        else:
            results = get_contact_by_name(query, include_all_matches=getattr(args, 'all_matches', False))

        if not results:
            print(f"No contact found for: {query}")
            return 0

        print(f"\n{'='*50}")
        print(f" Contact Results ({len(results)} match{'es' if len(results) > 1 else ''})")
        print(f"{'='*50}")

        for i, contact in enumerate(results, 1):
            if len(results) > 1:
                print(f"\n--- [{i}] ---")
            else:
                print()
            print(f"Display Name: {contact['display_name']}")
            print(f"Email: {contact['email']}")
            if contact.get('alias'):
                print(f"Alias: {contact['alias']}")
            if contact.get('first_name'):
                print(f"First Name: {contact['first_name']}")
            if contact.get('last_name'):
                print(f"Last Name: {contact['last_name']}")
            if contact.get('company'):
                print(f"Company: {contact['company']}")
            if contact.get('department'):
                print(f"Department: {contact['department']}")
            if contact.get('job_title'):
                print(f"Job Title: {contact['job_title']}")
            if contact.get('office'):
                print(f"Office: {contact['office']}")
            if contact.get('business_phone'):
                print(f"Business Phone: {contact['business_phone']}")
            if contact.get('mobile_phone'):
                print(f"Mobile: {contact['mobile_phone']}")
            if contact.get('city') or contact.get('state'):
                location_parts = [p for p in [contact.get('city'), contact.get('state')] if p]
                print(f"Location: {', '.join(location_parts)}")
            if contact.get('manager'):
                mgr = contact['manager']
                mgr_parts = [mgr.get('display_name', '')]
                if mgr.get('email'):
                    mgr_parts.append(f"<{mgr['email']}>")
                if mgr.get('job_title'):
                    mgr_parts.append(f"— {mgr['job_title']}")
                print(f"Manager: {' '.join(mgr_parts)}")

        return 0
    except Exception as e:
        print(f"Error: {str(e)}", file=sys.stderr)
        return 1

def cmd_find_thread(args):
    """Find all emails in the same conversation thread."""
    try:
        emails, message = find_thread_by_email_id(
            args.email_id,
            folder_names=(
                [f.strip() for f in args.folders.split(',')]
                if args.folders else None
            ),
            fuzzy=getattr(args, 'fuzzy', False),
        )
        print(f"\n\U0001F9F5 {message}\n")

        if emails:
            _print_thread_summary(emails)
            if getattr(args, 'brief', False):
                _display_email_list_brief(emails, show_folder=True)
            else:
                _display_email_list(emails, show_folder=True)

        return 0
    except Exception as e:
        print(f"Error: {str(e)}", file=sys.stderr)
        return 1

def cmd_download_attachment(args):
    """Download attachments from an email to a local directory."""
    try:
        with OutlookSessionManager() as session:
            email_item = _get_email_item(session, args.email_id)

            if email_item.Attachments.Count == 0:
                print("No attachments found on this email.")
                return 0

            save_dir = args.output_dir or os.path.join(os.path.expanduser("~"), "Downloads")
            os.makedirs(save_dir, exist_ok=True)

            saved = []
            for i in range(1, email_item.Attachments.Count + 1):
                attach = email_item.Attachments.Item(i)
                # Skip embedded images (cid: references) unless --all flag
                # Only skip if it looks like an image (not PDF/doc/etc)
                if not args.all and hasattr(attach, 'PropertyAccessor'):
                    try:
                        content_id = attach.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F")
                        if content_id:
                            ext = os.path.splitext(attach.FileName)[1].lower()
                            image_exts = {'.png', '.jpg', '.jpeg', '.gif', '.bmp', '.ico', '.svg'}
                            if ext in image_exts:
                                continue
                    except Exception:
                        pass

                if args.filename and attach.FileName.lower() != args.filename.lower():
                    continue

                save_path = os.path.join(save_dir, attach.FileName)
                attach.SaveAsFile(save_path)
                size_kb = os.path.getsize(save_path) / 1024
                saved.append((attach.FileName, save_path, size_kb))
                print(f"✅ Saved: {attach.FileName} ({size_kb:.1f} KB) → {save_path}")

            if not saved:
                if args.filename:
                    print(f"Attachment '{args.filename}' not found.")
                else:
                    print("No file attachments found (only embedded images).")
            else:
                print(f"\n{len(saved)} attachment(s) saved to: {save_dir}")

        return 0
    except Exception as e:
        print(f"Error: {str(e)}", file=sys.stderr)
        return 1

def cmd_find_related(args):
    """Find emails related to a given email using multiple strategies."""
    try:
        strategies = None
        if args.strategies:
            strategies = [s.strip() for s in args.strategies.split(',')]

        emails, message = find_related_emails(
            args.email_id,
            days=args.days,
            strategies=strategies,
            exclude_thread=args.exclude_thread,
            max_results=getattr(args, 'max_results', None),
        )
        print(f"\n\U0001F517 {message}\n")

        if emails:
            if getattr(args, 'brief', False):
                _display_email_list_brief(emails, show_folder=True)
            else:
                _display_email_list(emails, show_folder=True)

        return 0
    except Exception as e:
        print(f"Error: {str(e)}", file=sys.stderr)
        return 1

def cmd_get_html(args):
    """Get the raw HTMLBody of an email for template editing."""
    try:
        with OutlookSessionManager() as session:
            email_item = _get_email_item(session, args.email_id)

            subject = getattr(email_item, 'Subject', 'No Subject')
            sender = getattr(email_item, 'SenderName', 'Unknown')
            html_body = getattr(email_item, 'HTMLBody', '')

            print(f"Subject: {subject}")
            print(f"From: {sender}")
            print(f"To: {getattr(email_item, 'To', '')}")
            if getattr(email_item, 'CC', ''):
                print(f"CC: {email_item.CC}")
            print(f"HTML length: {len(html_body)} chars")
            print(f"\n{'='*60}")
            print(f"HTML_START")
            print(html_body)
            print(f"HTML_END")
            print(f"{'='*60}")

        return 0
    except Exception as e:
        print(f"Error: {str(e)}", file=sys.stderr)
        return 1
