"""HTML template editing commands."""

from __future__ import annotations

import os
import sys
import tempfile

from backend.outlook_session.session_manager import OutlookSessionManager

from .common import _get_email_item, _get_smtp_address, _normalize_cc


def cmd_edit_html(args):
    """Edit an email's HTML and save.

    If the email is a draft (unsent): modifies it IN PLACE and saves.
    If the email is sent/received: creates a new draft with modified content.
    """
    try:
        with OutlookSessionManager() as session:
            source_item = _get_email_item(session, args.email_id)

            is_draft = not getattr(source_item, 'Sent', True)

            # Start with source HTML
            html_body = getattr(source_item, 'HTMLBody', '')
            if not html_body:
                print("Error: Source email has no HTML body.", file=sys.stderr)
                return 1

            # Apply --body-file FIRST (full body replacement)
            if args.body_file:
                with open(args.body_file, 'r', encoding='utf-8') as f:
                    html_body = f.read()
                print(f"Body replaced from file: {args.body_file}")

            # Apply --replace pairs on top
            replacements_made = 0
            if args.replace:
                for pair in args.replace:
                    if '::' not in pair:
                        print(f"Error: Invalid replace format '{pair}'. Use 'old::new'.", file=sys.stderr)
                        return 1
                    old, new = pair.split('::', 1)
                    if old in html_body:
                        html_body = html_body.replace(old, new)
                        replacements_made += 1
                    else:
                        print(f"WARNING: '{old}' not found in HTML body.")

            if is_draft:
                # --- Modify existing draft in place ---
                draft = source_item

                if args.subject:
                    draft.Subject = args.subject
                if args.to:
                    draft.To = args.to.replace(',', '; ')
                cc_normalized = _normalize_cc(args.cc)
                if cc_normalized:
                    draft.CC = cc_normalized.replace(',', '; ')

                draft.HTMLBody = html_body
                draft.Save()

                print(f"✅ Draft updated in place")
            else:
                # --- Create new draft from sent/received email ---
                draft = session.outlook.CreateItem(0)  # olMailItem

                if args.subject:
                    draft.Subject = args.subject
                else:
                    draft.Subject = getattr(source_item, 'Subject', '')

                if args.to:
                    draft.To = args.to.replace(',', '; ')
                else:
                    to_addrs = []
                    for i in range(1, source_item.Recipients.Count + 1):
                        recip = source_item.Recipients.Item(i)
                        if recip.Type == 1:  # olTo
                            to_addrs.append(_get_smtp_address(recip))
                    draft.To = '; '.join(to_addrs)

                cc_normalized = _normalize_cc(args.cc)
                if cc_normalized:
                    draft.CC = cc_normalized.replace(',', '; ')
                else:
                    cc_addrs = []
                    for i in range(1, source_item.Recipients.Count + 1):
                        recip = source_item.Recipients.Item(i)
                        if recip.Type == 2:  # olCC
                            cc_addrs.append(_get_smtp_address(recip))
                    if cc_addrs:
                        draft.CC = '; '.join(cc_addrs)

                # Set modified HTML body first
                draft.HTMLBody = html_body

                # Copy embedded images to preserve cid: references in HTML
                import tempfile
                for i in range(1, source_item.Attachments.Count + 1):
                    attach = source_item.Attachments.Item(i)
                    try:
                        cid = attach.PropertyAccessor.GetProperty(
                            "http://schemas.microsoft.com/mapi/proptag/0x3712001F"
                        )
                    except Exception:
                        cid = None
                    if cid:
                        temp_path = os.path.join(tempfile.gettempdir(), attach.FileName)
                        attach.SaveAsFile(temp_path)
                        new_attach = draft.Attachments.Add(temp_path)
                        new_attach.PropertyAccessor.SetProperty(
                            "http://schemas.microsoft.com/mapi/proptag/0x3712001F", cid
                        )
                        os.remove(temp_path)

                # Force Outlook to re-resolve CID references
                draft.HTMLBody = draft.HTMLBody

                # Copy file attachments (non-embedded) from source if requested
                if args.copy_attachments:
                    for i in range(1, source_item.Attachments.Count + 1):
                        attach = source_item.Attachments.Item(i)
                        try:
                            cid = attach.PropertyAccessor.GetProperty(
                                "http://schemas.microsoft.com/mapi/proptag/0x3712001F"
                            )
                            if cid:
                                continue
                        except Exception:
                            pass
                        import tempfile
                        temp_path = os.path.join(tempfile.gettempdir(), attach.FileName)
                        attach.SaveAsFile(temp_path)
                        draft.Attachments.Add(temp_path)

                draft.Save()
                print(f"✅ New draft created in Drafts folder")

            print(f"   EntryID: {draft.EntryID}")
            print(f"   Subject: {draft.Subject}")
            print(f"   To: {draft.To}")
            if draft.CC:
                print(f"   CC: {draft.CC}")
            if replacements_made:
                print(f"   Replacements applied: {replacements_made}")
            print(f"   HTML length: {len(html_body)} chars")

        return 0
    except Exception as e:
        print(f"Error: {str(e)}", file=sys.stderr)
        return 1
