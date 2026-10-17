"""Commands that create or send Outlook mail."""

from __future__ import annotations

import base64
from datetime import timedelta
import os
import sys
import time

from backend.config import batch_config
from backend.outlook_session.session_manager import OutlookSessionManager

from .common import (
    _add_attachments,
    _add_inline_images,
    _add_recipients,
    _attach_emails,
    _bare_subject,
    _build_reply_header,
    _collect_recent_sent_entry_ids,
    _ensure_utf8_charset,
    _format_forward_message_html,
    _get_email_item,
    _get_smtp_address,
    _inject_before_body,
    _normalize_cc,
    _preserve_inline_images,
    _print_sent_entry_id,
    _remove_self_from_recipients,
    _reply_step,
    _resolve_body,
    _set_importance,
    _wrap_body_font,
)


def cmd_reply(args):
    """Reply to an email. Default: reply-all. --only: sender (From) only."""
    if not _resolve_body(args):
        return 1
    try:
        _reply_step("opening Outlook session")
        with OutlookSessionManager() as session:
            _reply_step("loading source email")
            email_item = _get_email_item(session, args.email_id)
            _reply_step("source email loaded")
            current_user = session.outlook.Session.CurrentUser
            current_user_email = (
                current_user.AddressEntry.GetExchangeUser().PrimarySmtpAddress
                if current_user.AddressEntry.GetExchangeUser() else ""
            )
            _reply_step("current user resolved")

            parent_folder = email_item.Parent
            is_sent_items = "Sent Items" in parent_folder.Name or "已发送邮件" in parent_folder.Name

            if args.only:
                # --only: reply to From (sender) only
                if is_sent_items:
                    _reply_step("creating new reply item from Sent Items source")
                    reply = session.outlook.CreateItem(0)
                    _reply_step("adding explicit recipients")
                    _add_recipients(reply, args.to, args.cc)
                    _reply_step("displaying reply item to load signature")
                    reply.Display(False)
                    signature_html = reply.HTMLBody
                    original_body = email_item.HTMLBody if email_item.HTMLBody else f"<p>{email_item.Body}</p>"
                    _reply_step("setting HTML body")
                    reply.HTMLBody = _wrap_body_font(args.body) + signature_html + _build_reply_header(email_item, current_user_email) + original_body
                else:
                    _reply_step("creating sender-only reply item")
                    reply = email_item.Reply()
                    _reply_step("adding explicit recipients")
                    _add_recipients(reply, args.to, args.cc)
                    _reply_step("setting HTML body")
                    reply.HTMLBody = _wrap_body_font(args.body) + reply.HTMLBody
            elif is_sent_items:
                # Sent Items: create new email with original recipients
                _reply_step("creating new reply-all item from Sent Items source")
                reply = session.outlook.CreateItem(0)
                _reply_step("copying original recipients")
                for i in range(1, email_item.Recipients.Count + 1):
                    recip = email_item.Recipients.Item(i)
                    smtp = _get_smtp_address(recip)
                    if smtp and smtp.lower() != current_user_email.lower():
                        new_recip = reply.Recipients.Add(smtp)
                        new_recip.Type = recip.Type  # 1=To, 2=CC
                _reply_step("adding explicit recipients")
                _add_recipients(reply, args.to, args.cc)
                _reply_step("displaying reply item to load signature")
                reply.Display(False)
                signature_html = reply.HTMLBody
                original_body = email_item.HTMLBody if email_item.HTMLBody else f"<p>{email_item.Body}</p>"
                _reply_step("setting HTML body")
                reply.HTMLBody = _wrap_body_font(args.body) + signature_html + _build_reply_header(email_item, current_user_email) + original_body
            else:
                # Default: ReplyAll
                _reply_step("creating reply-all item")
                reply = email_item.ReplyAll()
                _reply_step("removing current user from recipients")
                _remove_self_from_recipients(reply, current_user_email)
                _reply_step("adding explicit recipients")
                _add_recipients(reply, args.to, args.cc)
                _reply_step("setting HTML body")
                reply.HTMLBody = _wrap_body_font(args.body) + reply.HTMLBody

            _reply_step("ensuring UTF-8 charset")
            reply.HTMLBody = _ensure_utf8_charset(reply.HTMLBody)
            _reply_step("preserving inline images")
            _preserve_inline_images(reply, email_item)
            _reply_step("setting subject")
            reply.Subject = f"RE: {email_item.Subject}" if not email_item.Subject.startswith("RE:") else email_item.Subject

            _reply_step("checking recipients")
            if reply.Recipients.Count == 0:
                print("Error: No recipients found.")
                return 1

            count = reply.Recipients.Count
            _reply_step("adding inline images")
            _add_inline_images(reply, args.inline_image)
            _reply_step("adding attachments")
            _add_attachments(reply, args.attach)
            _reply_step("attaching emails")
            _attach_emails(reply, args.attach_email, session)
            _reply_step("setting importance")
            _set_importance(reply, args.importance)
            subject = reply.Subject
            sent_before_ids = _collect_recent_sent_entry_ids(session, subject)
            from datetime import datetime, timezone
            send_start_time = datetime.now(timezone.utc)
            _reply_step("sending")
            reply.Send()
            _reply_step("send returned")
            mode = "Reply (From only)" if args.only else "Reply-all"
            print(f"{mode} sent to {count} recipient(s)")
            _reply_step("locating sent EntryID")
            _print_sent_entry_id(session, subject, sent_after=send_start_time, exclude_entry_ids=sent_before_ids)
            _reply_step("sent EntryID lookup complete")
            return 0
    except Exception as e:
        print(f"Error: {str(e)}", file=sys.stderr)
        return 1

def cmd_compose(args):
    """Compose and send new email (always HTML format)"""
    if not _resolve_body(args):
        return 1
    try:
        to_list = [x.strip() for x in args.to.replace(";", ",").split(",")] if args.to else []
        cc_normalized = _normalize_cc(args.cc)
        cc_list = [x.strip() for x in cc_normalized.replace(";", ",").split(",")] if cc_normalized else []

        with OutlookSessionManager() as session:
            mail = session.outlook.CreateItem(0)  # 0 = olMailItem

            # Set recipients
            for recipient in to_list:
                mail.Recipients.Add(recipient)
            if cc_list:
                for cc_recipient in cc_list:
                    cc_recip = mail.Recipients.Add(cc_recipient)
                    cc_recip.Type = 2  # 2 = olCC

            mail.Subject = args.subject

            # Display briefly to trigger Outlook signature insertion
            mail.Display(False)

            # Prepend body to signature HTML (same pattern as reply)
            mail.HTMLBody = _ensure_utf8_charset(_wrap_body_font(args.body) + mail.HTMLBody)

            _add_inline_images(mail, args.inline_image)
            _add_attachments(mail, args.attach)
            _attach_emails(mail, args.attach_email, session)
            _set_importance(mail, args.importance)
            sent_before_ids = _collect_recent_sent_entry_ids(session, args.subject)
            from datetime import datetime, timezone
            send_start_time = datetime.now(timezone.utc)
            mail.Send()
            total_recipients = len(to_list) + len(cc_list)
            print(f"HTML email sent successfully to {total_recipients} recipient(s)")
            _print_sent_entry_id(session, args.subject, sent_after=send_start_time, exclude_entry_ids=sent_before_ids)

        return 0
    except Exception as e:
        print(f"Error: {str(e)}", file=sys.stderr)
        return 1

def cmd_batch_forward(args):
    """Batch forward email to multiple recipients by email ID"""
    if getattr(args, 'body_file', None):
        try:
            with open(args.body_file, 'r', encoding='utf-8') as f:
                args.message = f.read()
        except OSError as exc:
            print(f"Error reading body file {args.body_file}: {exc}", file=sys.stderr)
            return 1
    elif getattr(args, 'message_base64', None):
        try:
            args.message = base64.b64decode(args.message_base64).decode('utf-8')
        except Exception as exc:
            print(f"Error: invalid --body-base64/--message-base64 payload: {exc}", file=sys.stderr)
            return 1

    try:
        # Use configured batch size from backend/config.py
        batch_size = batch_config.OUTLOOK_BCC_LIMIT
        
        with OutlookSessionManager() as session:
            email_item = _get_email_item(session, args.email_id)
            email_subject = str(getattr(email_item, "Subject", "No Subject"))

            # Read CSV file (handle BOM if present)
            import csv
            recipients = []
            with open(args.csv_path, 'r', encoding='utf-8-sig') as f:
                reader = csv.DictReader(f)
                for row in reader:
                    row_lower = {k.lower(): v for k, v in row.items()}
                    if 'email' in row_lower:
                        recipients.append(row_lower['email'].strip())
            
            if not recipients:
                print("Error: No email addresses found in CSV", file=sys.stderr)
                return 1
            
            # Forward to recipients in batches (batch size from config file)
            import time
            from datetime import datetime, timezone
            total_sent = 0
            num_batches = (len(recipients) + batch_size - 1) // batch_size
            fw_subject = f"FW: {email_subject}" if not email_subject.startswith("FW:") else email_subject
            existing_sent_ids = _collect_recent_sent_entry_ids(session, fw_subject)
            send_start_time = datetime.now(timezone.utc)

            for i in range(0, len(recipients), batch_size):
                batch = recipients[i:i + batch_size]
                batch_num = i // batch_size + 1

                # Create forward
                forward = email_item.Forward()

                # Add custom message if provided (insert into body tag like reply does)
                if args.message:
                    forward.HTMLBody = _format_forward_message_html(args.message) + forward.HTMLBody

                # Add recipients as BCC (to protect privacy)
                for recipient in batch:
                    bcc_recip = forward.Recipients.Add(recipient)
                    bcc_recip.Type = 3  # 3 = olBCC

                # Resolve all recipients before sending
                forward.Recipients.ResolveAll()

                # Send
                forward.Send()
                total_sent += len(batch)
                print(f"Sent batch {batch_num}/{num_batches}: {len(batch)} recipients")

            print(f"\nTotal recipients: {total_sent}")
            print(f"Successfully forwarded email: {email_subject}")

            # Retrieve EntryIDs for all sent batches
            bare_subject = _bare_subject(fw_subject)
            threshold = send_start_time - timedelta(seconds=5)

            for attempt in range(15):
                time.sleep(2)
                try:
                    ns = session.outlook.GetNamespace("MAPI")
                    sent_folder = ns.GetDefaultFolder(5)
                    items = sent_folder.Items
                    items.Sort("[SentOn]", True)
                    found_ids = []
                    item = items.GetFirst()
                    while item and len(found_ids) < num_batches:
                        sent_on = item.SentOn
                        if hasattr(sent_on, 'astimezone'):
                            sent_utc = sent_on.astimezone(timezone.utc)
                        else:
                            sent_utc = sent_on.replace(tzinfo=timezone.utc)
                        if sent_utc < threshold:
                            break
                        bare_item_subject = _bare_subject(getattr(item, 'Subject', ''))
                        entry_id = getattr(item, 'EntryID', None)
                        if entry_id in existing_sent_ids:
                            item = items.GetNext()
                            continue
                        if bare_subject in bare_item_subject and entry_id:
                            found_ids.append(entry_id)
                        item = items.GetNext()
                    if len(found_ids) >= num_batches:
                        for idx, eid in enumerate(reversed(found_ids), 1):
                            print(f"EntryID (batch {idx}): {eid}")
                        break
                except Exception:
                    pass

        return 0
    except Exception as e:
        print(f"Error: {str(e)}", file=sys.stderr)
        return 1

def cmd_forward(args):
    """Forward an email to specified recipients with optional CC and custom message."""
    if not _resolve_body(args, required=False):
        return 1
    try:
        with OutlookSessionManager() as session:
            email_item = _get_email_item(session, args.email_id)

            forward = email_item.Forward()

            # Meeting items preserve original attendees as recipients — clear them
            while forward.Recipients.Count > 0:
                forward.Recipients.Remove(1)

            # Remove attachments if requested
            if getattr(args, 'no_attachments', False):
                while forward.Attachments.Count > 0:
                    forward.Attachments.Remove(1)

            # Set subject: use --subject override or default FW: prefix
            original_subject = str(getattr(email_item, "Subject", "No Subject"))
            if getattr(args, 'subject', None):
                forward.Subject = args.subject
            else:
                forward.Subject = f"FW: {original_subject}" if not original_subject.startswith("FW:") else original_subject

            # Add To recipients
            if args.to:
                for r in args.to.replace(";", ",").split(","):
                    r = r.strip()
                    if r:
                        forward.Recipients.Add(r)

            # Add CC recipients (args.cc is a list from action='append', each entry may be comma-separated)
            if args.cc:
                cc_joined = ",".join(args.cc)
                for r in cc_joined.replace(";", ",").split(","):
                    r = r.strip()
                    if r:
                        cc_recip = forward.Recipients.Add(r)
                        cc_recip.Type = 2  # 2 = olCC

            # Prepend custom message if provided
            if args.body:
                prepend_html = _format_forward_message_html(args.body)
                try:
                    original_html = forward.HTMLBody
                    forward.HTMLBody = _inject_before_body(prepend_html, original_html)
                except Exception:
                    # Meeting items / special types need inspector initialization
                    _ = forward.GetInspector
                    forward.HTMLBody = _inject_before_body(prepend_html, forward.HTMLBody)

            if forward.Recipients.Count == 0:
                print("Error: No recipients specified. Use --to.")
                return 1

            resolved = forward.Recipients.ResolveAll()
            if resolved is False:
                print("Error: One or more recipients could not be resolved.")
                return 1

            recipient_count = forward.Recipients.Count
            final_subject = str(getattr(forward, "Subject", original_subject))

            _add_inline_images(forward, args.inline_image)
            _add_attachments(forward, args.attach)
            _attach_emails(forward, args.attach_email, session)
            _set_importance(forward, args.importance)
            sent_before_ids = _collect_recent_sent_entry_ids(session, final_subject)
            from datetime import datetime, timezone
            send_start_time = datetime.now(timezone.utc)
            forward.Send()
            print(f"Forward sent to {recipient_count} recipient(s)")
            print(f"Subject: {final_subject}")
            _print_sent_entry_id(session, final_subject, sent_after=send_start_time, exclude_entry_ids=sent_before_ids)
            return 0
    except Exception as e:
        print(f"Error: {str(e)}", file=sys.stderr)
        return 1

def cmd_redirect(args):
    """Redirect an email: clear all recipients, add new TO/CC, preserve original body."""
    if not _resolve_body(args):
        return 1
    try:
        with OutlookSessionManager() as session:
            email_item = _get_email_item(session, args.email_id)
            forward = email_item.Forward()

            # Remove all pre-populated recipients
            while forward.Recipients.Count > 0:
                forward.Recipients.Remove(1)

            # Add new recipients
            _add_recipients(forward, args.to, args.cc)

            if forward.Recipients.Count == 0:
                print("Error: No recipients specified. Use --to or --cc.")
                return 1

            # Subject
            original_subject = str(getattr(email_item, "Subject", "No Subject"))
            forward.Subject = f"FW: {original_subject}" if not original_subject.startswith("FW:") else original_subject

            # Signature + body
            forward.Display(False)
            signature_html = forward.HTMLBody
            forward.HTMLBody = _ensure_utf8_charset(_wrap_body_font(args.body) + signature_html)

            recipient_count = forward.Recipients.Count
            final_subject = forward.Subject

            _add_inline_images(forward, args.inline_image)
            _add_attachments(forward, args.attach)
            sent_before_ids = _collect_recent_sent_entry_ids(session, final_subject)
            from datetime import datetime, timezone
            send_start_time = datetime.now(timezone.utc)
            forward.Send()
            print(f"Redirected to {recipient_count} recipient(s)")
            print(f"Subject: {final_subject}")
            _print_sent_entry_id(session, final_subject, sent_after=send_start_time, exclude_entry_ids=sent_before_ids)
            return 0
    except Exception as e:
        print(f"Error: {str(e)}", file=sys.stderr)
        return 1

def cmd_send_draft(args):
    """Send an existing draft email."""
    try:
        with OutlookSessionManager() as session:
            email_item = _get_email_item(session, args.email_id)

            if getattr(email_item, 'Sent', False):
                print("Error: This email has already been sent.", file=sys.stderr)
                return 1

            subject = getattr(email_item, 'Subject', 'No Subject')
            to = getattr(email_item, 'To', '')
            cc = getattr(email_item, 'CC', '')

            if not to:
                print("Error: Draft has no recipients.", file=sys.stderr)
                return 1

            # Resolve all recipients before sending
            print("Resolving recipients...")
            resolved = email_item.Recipients.ResolveAll()
            if not resolved:
                print("Warning: Some recipients could not be automatically resolved. Checking status...", file=sys.stderr)
                unresolved = []
                for i in range(1, email_item.Recipients.Count + 1):
                    recip = email_item.Recipients.Item(i)
                    if not recip.Resolved:
                        unresolved.append(recip.Name)
                if unresolved:
                    print(f"Error: Unresolved recipients found: {', '.join(unresolved)}", file=sys.stderr)
                    print("Please open Outlook to manually resolve these names in the Drafts folder.", file=sys.stderr)
                    return 1

            try:
                email_item.Send()
            except Exception:
                email_item.Display(False)
                import time
                time.sleep(0.5)
                email_item.Send()
            print(f"✅ Email sent")
            print(f"   Subject: {subject}")
            print(f"   To: {to}")
            if cc:
                print(f"   CC: {cc}")

        return 0
    except Exception as e:
        print(f"Error: {str(e)}", file=sys.stderr)
        return 1
