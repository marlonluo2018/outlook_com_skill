"""Argument parser for the Outlook COM skill CLI."""

from __future__ import annotations

import argparse

from backend.config import search_config

from .calendar_commands import (
    cmd_cancel_invite,
    cmd_check_availability,
    cmd_delete_appointment,
    cmd_list_calendar,
    cmd_respond_meeting,
)
from .folder_commands import (
    cmd_create_folder,
    cmd_delete_email,
    cmd_move_email,
    cmd_recall,
    cmd_remove_folder,
)
from .mail_commands import (
    cmd_batch_compose,
    cmd_batch_forward,
    cmd_compose,
    cmd_forward,
    cmd_redirect,
    cmd_reply,
    cmd_send_draft,
)
from .search_commands import (
    cmd_download_attachment,
    cmd_find_related,
    cmd_find_thread,
    cmd_get_email,
    cmd_get_html,
    cmd_list_folders,
    cmd_list_recent,
    cmd_lookup_contact,
    cmd_search,
)
from .settings_commands import (
    cmd_disable_ooo,
    cmd_get_ooo,
    cmd_get_signature,
    cmd_set_ooo,
    cmd_update_signature,
)
from .template_commands import cmd_edit_html


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description='Outlook Skill - Email Management CLI')
    subparsers = parser.add_subparsers(dest='command', help='Available commands')
    
    # List folders command
    parser_list_folders = subparsers.add_parser('list-folders', help='List all Outlook folders')
    parser_list_folders.add_argument('--hide-system', action='store_true', help='Hide system folders')
    parser_list_folders.set_defaults(func=cmd_list_folders)
    
    # List recent emails command
    parser_list_recent = subparsers.add_parser('find-recent', help='Find recent emails with IDs')
    parser_list_recent.add_argument('--days', type=int, default=7, help='Days back to search (1-30)')
    parser_list_recent.add_argument('--folder', type=str, default=None, help='Folder name (default: Inbox)')
    parser_list_recent.add_argument('--limit', type=int, default=None, help='Max number of results to return (default: all)')
    parser_list_recent.add_argument('--json', action='store_true', help='Output stable JSON including EntryID and ConversationID')
    parser_list_recent.set_defaults(func=cmd_list_recent)
    
    # Search emails command
    parser_search = subparsers.add_parser('find', help='Find emails by sender, subject, recipient, or body')
    parser_search.add_argument('--from', dest='from_', type=str, default=None, help='Search by sender name/email')
    parser_search.add_argument('--subject', type=str, default=None, help='Search by subject')
    parser_search.add_argument('--to', type=str, default=None, help='Search by recipient name/email')
    parser_search.add_argument('--body', type=str, default=None, help='Search by body content')
    parser_search.add_argument('--type', type=str, default=None, choices=['subject', 'sender', 'recipient', 'body'], help=argparse.SUPPRESS)
    parser_search.add_argument('--query', type=str, default=None, help=argparse.SUPPRESS)
    parser_search.add_argument(
        '--days',
        type=int,
        default=search_config.DIRECT_FIND_DEFAULT_DAYS,
        help=(
            f"Days back to search "
            f"(default: {search_config.DIRECT_FIND_DEFAULT_DAYS}, "
            f"allowed range: 1-{search_config.MAX_SEARCH_DAYS})"
        ),
    )
    parser_search.add_argument('--folder', type=str, default=None, help='Folder name (default depends on field)')
    parser_search.add_argument('--folders', type=str, default=None, help='Comma-separated folder names for cross-folder search')
    parser_search.add_argument('--limit', type=int, default=None, help='Max number of results to return (default: all)')
    parser_search.add_argument('--match-all', action='store_true', default=True, help='Match all terms (AND logic)')
    parser_search.set_defaults(func=cmd_search)
    
    # Get email command
    parser_get_email = subparsers.add_parser('get-email', aliases=['read'], help='Get full email details by ID')
    parser_get_email.add_argument('email_ids', nargs='*', default=[], help='One or more Email IDs from search results')
    parser_get_email.add_argument('--id', dest='email_id_flag', help='Email ID (alternative to positional)')
    parser_get_email.add_argument('--subject', help='Exact subject from find output — used to locate the email when the EntryID cannot be opened in this Outlook session')
    parser_get_email.add_argument('--pick', type=int, help='Row number from the last find result set (1-based) — avoids copying an EntryID')
    parser_get_email.add_argument('--last', action='store_true', help='Use the last row of the last find result set')
    parser_get_email.add_argument('--task', help='Task ID or file (e.g. T152) — resolves the email marker recorded in that task file')
    parser_get_email.add_argument('--fields', help='Comma-separated fields to show (from,to,cc,subject,date,body)')
    parser_get_email.add_argument('--truncate', type=int, default=None, help='Truncate the email body to N characters')
    parser_get_email.add_argument('--latest-only', '-l', action='store_true', help='Show only the latest message body, excluding quoted thread history')
    parser_get_email.add_argument('--max-lines', type=int, default=None, help='Show up to N lines of the email body')
    parser_get_email.set_defaults(func=cmd_get_email)
    
    # Reply command (default: reply-all; --only: sender only)
    parser_reply = subparsers.add_parser('reply', help='Reply to email (default: reply-all; --only: From only)')
    parser_reply.add_argument('email_id', help='Email ID from search results')
    parser_reply.add_argument('body', nargs='?', default=None, help='Reply text in HTML format')
    parser_reply.add_argument('--body', dest='body', help='Reply text in HTML format')
    parser_reply.add_argument('--body-stdin', action='store_true', help='Read body from stdin pipe')
    parser_reply.add_argument('--body-stdin-base64', action='store_true', help='Read base64-encoded UTF-8 HTML body from stdin')
    parser_reply.add_argument('--body-base64', help='Base64 encoded UTF-8 HTML body')
    parser_reply.add_argument('--body-file', help='Path to file containing the HTML body')
    parser_reply.add_argument('--to', help='Additional To recipients (comma separated)')
    parser_reply.add_argument('--cc', action='append', help='CC recipients (comma separated or repeated)')
    parser_reply.add_argument('--attach', help='File path(s) to attach (comma separated)')
    parser_reply.add_argument('--attach-email', help='Email ID(s) to attach as .msg (comma separated)')
    parser_reply.add_argument('--inline-image', help='Image path(s) to embed inline in body (comma separated)')
    parser_reply.add_argument('--importance', choices=['high', 'low'], help='Set email importance (high or low)')
    parser_reply.add_argument('--only', action='store_true', help='Reply to From (sender) only instead of reply-all')
    parser_reply.set_defaults(func=cmd_reply)
    
    # Compose command
    parser_compose = subparsers.add_parser('compose', help='Compose and send new email')
    parser_compose.add_argument('--to', required=True, help='To recipients (comma separated)')
    parser_compose.add_argument('--subject', required=True, help='Email subject')
    parser_compose.add_argument('--body', help='Email body')
    parser_compose.add_argument('--body-stdin', action='store_true', help='Read body from stdin pipe')
    parser_compose.add_argument('--body-stdin-base64', action='store_true', help='Read base64-encoded UTF-8 HTML body from stdin')
    parser_compose.add_argument('--body-base64', help='Base64 encoded UTF-8 HTML body')
    parser_compose.add_argument('--body-file', help='Path to file containing the HTML body')
    parser_compose.add_argument('--cc', action='append', help='CC recipients (comma separated or repeated)')
    parser_compose.add_argument('--attach', help='File path(s) to attach (comma separated)')
    parser_compose.add_argument('--attach-email', help='Email ID(s) to attach as .msg (comma separated)')
    parser_compose.add_argument('--inline-image', help='Image path(s) to embed inline in body (comma separated)')
    parser_compose.add_argument('--importance', choices=['high', 'low'], help='Set email importance (high or low)')
    parser_compose.set_defaults(func=cmd_compose)
    
    # Forward command
    parser_forward = subparsers.add_parser('forward', help='Forward an email to specified recipients')
    parser_forward.add_argument('email_id', help='Email ID from search results')
    parser_forward.add_argument('--to', required=True, help='To recipients (comma separated)')
    parser_forward.add_argument('--cc', action='append', help='CC recipients (comma separated or repeated)')
    parser_forward.add_argument('--body', help='Custom message to prepend')
    parser_forward.add_argument('--body-stdin', action='store_true', help='Read body from stdin pipe')
    parser_forward.add_argument('--body-stdin-base64', action='store_true', help='Read base64-encoded UTF-8 HTML body from stdin')
    parser_forward.add_argument('--body-base64', help='Base64 encoded UTF-8 HTML body')
    parser_forward.add_argument('--body-file', help='Path to file containing the HTML body')
    parser_forward.add_argument('--attach', help='File path(s) to attach (comma separated)')
    parser_forward.add_argument('--attach-email', help='Email ID(s) to attach as .msg (comma separated)')
    parser_forward.add_argument('--inline-image', help='Image path(s) to embed inline in body (comma separated)')
    parser_forward.add_argument('--importance', choices=['high', 'low'], help='Set email importance (high or low)')
    parser_forward.add_argument('--subject', help='Override the forwarded subject (default: FW: <original>)')
    parser_forward.add_argument('--no-attachments', action='store_true', help='Remove any attachments from the forwarded email')
    parser_forward.set_defaults(func=cmd_forward)

    # Redirect command (clear all recipients, add new ones)
    parser_redirect = subparsers.add_parser('redirect', help='Redirect email: clear all recipients, set new TO/CC')
    parser_redirect.add_argument('email_id', help='Email ID from search results')
    parser_redirect.add_argument('body', nargs='?', default=None, help='Message body in HTML format')
    parser_redirect.add_argument('--body-stdin', action='store_true', help='Read body from stdin pipe')
    parser_redirect.add_argument('--body-stdin-base64', action='store_true', help='Read base64-encoded UTF-8 HTML body from stdin')
    parser_redirect.add_argument('--body-base64', help='Base64 encoded UTF-8 HTML body')
    parser_redirect.add_argument('--body-file', help='Path to file containing the HTML body')
    parser_redirect.add_argument('--to', required=True, help='To recipients (comma separated)')
    parser_redirect.add_argument('--cc', action='append', help='CC recipients (comma separated or repeated)')
    parser_redirect.add_argument('--attach', help='File path(s) to attach (comma separated; duplicates of attachments inherited from the original email are skipped)')
    parser_redirect.add_argument('--inline-image', help='Image path(s) to embed inline in body (comma separated)')
    parser_redirect.add_argument('--subject', help='Override the subject (default: FW: <original>)')
    parser_redirect.add_argument('--no-attachments', action='store_true', help='Remove any attachments inherited from the original email')
    parser_redirect.set_defaults(func=cmd_redirect)

    # Batch forward command
    parser_batch = subparsers.add_parser('batch-forward', help='Batch forward email by ID to multiple recipients')
    parser_batch.add_argument('email_id', help='Email ID from search results')
    parser_batch.add_argument('csv_path', help='Path to CSV file with email addresses')
    parser_batch.add_argument('--message', help='Custom message to prepend (HTML format)')
    parser_batch.add_argument('--body-base64', dest='message_base64', help='Base64 encoded UTF-8 custom message to prepend (HTML format)')
    parser_batch.add_argument('--message-base64', dest='message_base64', help='Alias for --body-base64')
    parser_batch.add_argument('--body-file', help='Path to file containing the custom HTML message')
    parser_batch.set_defaults(func=cmd_batch_forward)

    # Batch compose command (new email to CSV recipients via BCC)
    parser_batch_compose = subparsers.add_parser('batch-compose', help='Compose and send a new email to many recipients (CSV) as BCC')
    parser_batch_compose.add_argument('csv_path', help='Path to CSV file with email addresses (column: email)')
    parser_batch_compose.add_argument('--subject', required=True, help='Email subject')
    parser_batch_compose.add_argument('--body-base64', dest='message_base64', help='Base64 encoded UTF-8 HTML body')
    parser_batch_compose.add_argument('--message-base64', dest='message_base64', help='Alias for --body-base64')
    parser_batch_compose.add_argument('--body-file', help='Path to file containing the HTML body')
    parser_batch_compose.add_argument('--cc', action='append', help='CC recipients (comma separated or repeated)')
    parser_batch_compose.add_argument('--importance', choices=['high', 'low'], help='Set email importance (high or low)')
    parser_batch_compose.set_defaults(func=cmd_batch_compose)
    
    # Create folder command
    parser_create_folder = subparsers.add_parser('create-folder', help='Create a new folder')
    parser_create_folder.add_argument('name', help='Folder name')
    parser_create_folder.add_argument('--parent', help='Parent folder name')
    parser_create_folder.set_defaults(func=cmd_create_folder)
    
    # Remove folder command
    parser_remove_folder = subparsers.add_parser('remove-folder', help='Remove a folder')
    parser_remove_folder.add_argument('name', help='Folder name or path')
    parser_remove_folder.set_defaults(func=cmd_remove_folder)
    
    # Move email command
    parser_move = subparsers.add_parser('move-email', help='Move an email to a folder by ID')
    parser_move.add_argument('email_id', help='Email ID from search results')
    parser_move.add_argument('folder', help='Target folder name')
    parser_move.set_defaults(func=cmd_move_email)
    
    # Delete email command
    parser_delete = subparsers.add_parser('delete-email', help='Delete an email by ID')
    parser_delete.add_argument('email_id', help='Email ID from search results')
    parser_delete.set_defaults(func=cmd_delete_email)

    # Recall email command
    parser_recall = subparsers.add_parser('recall', help='Recall a sent email via Exchange')
    parser_recall.add_argument('email_id', help='Email ID from search results (must be in Sent Items)')
    parser_recall.set_defaults(func=cmd_recall)

    # Lookup contact command
    parser_lookup = subparsers.add_parser('lookup-contact', help='Look up contact information by email or display name')
    parser_lookup.add_argument('query', help='Email address or display name to look up')
    parser_lookup.add_argument('--all', dest='all_matches', action='store_true', help='Search GAL for all nearby matches; slower')
    parser_lookup.add_argument('--timeout', type=int, default=None, help='Max seconds to allow lookup before stopping')
    parser_lookup.set_defaults(func=cmd_lookup_contact)

    # Find thread command
    parser_thread = subparsers.add_parser('find-thread', help='Find all emails in same conversation thread')
    parser_thread.add_argument('email_id', help='Email ID from search results')
    parser_thread.add_argument('--folders', type=str, default=None, help='Folders to search (default: Inbox,Sent Items)')
    parser_thread.add_argument('--fuzzy', action='store_true', help='Also find emails with similar subjects (token overlap)')
    parser_thread.add_argument('--brief', action='store_true', help='Compact single-line output')
    parser_thread.set_defaults(func=cmd_find_thread)

    # Find related command
    parser_related = subparsers.add_parser('find-related', help='Find emails related to a given email')
    parser_related.add_argument('email_id', help='Email ID from search results')
    parser_related.add_argument('--days', type=int, default=90, help='Days back for sender/keyword strategies')
    parser_related.add_argument('--strategies', type=str, default=None, help='Strategies: thread,sender,recipient,keyword (default: all)')
    parser_related.add_argument('--exclude-thread', action='store_true', help='Skip thread strategy (useful after find-thread)')
    parser_related.add_argument('--limit', type=int, dest='max_results', default=None, help='Maximum results to return (default: from config)')
    parser_related.add_argument('--brief', action='store_true', help='Compact single-line output')
    parser_related.set_defaults(func=cmd_find_related)

    # Download attachment command
    parser_download = subparsers.add_parser('download-attachment', help='Download attachments from an email to local directory')
    parser_download.add_argument('email_id', help='Email ID from search results')
    parser_download.add_argument('--output-dir', help='Directory to save attachments (default: ~/Downloads)')
    parser_download.add_argument('--filename', help='Download only this specific attachment filename')
    parser_download.add_argument('--all', action='store_true', help='Include embedded images (default: skip them)')
    parser_download.set_defaults(func=cmd_download_attachment)

    # Get HTML command
    parser_get_html = subparsers.add_parser('get-html', help='Get raw HTMLBody of an email for template editing')
    parser_get_html.add_argument('email_id', help='Email ID from search results')
    parser_get_html.set_defaults(func=cmd_get_html)

    # Send draft command
    parser_send_draft = subparsers.add_parser('send-draft', help='Send an existing draft email')
    parser_send_draft.add_argument('email_id', help='Draft email ID')
    parser_send_draft.set_defaults(func=cmd_send_draft)

    # Edit HTML command (template editing → save to Drafts)
    parser_edit_html = subparsers.add_parser('edit-html', help='Edit email HTML and save as new draft')
    parser_edit_html.add_argument('email_id', help='Source email ID (used as template)')
    parser_edit_html.add_argument('--replace', action='append', help='Text replacement: "old::new" (repeatable)')
    parser_edit_html.add_argument('--subject', help='Override subject line')
    parser_edit_html.add_argument('--to', help='Override To recipients (comma separated)')
    parser_edit_html.add_argument('--cc', action='append', help='CC recipients (comma separated or repeated)')
    parser_edit_html.add_argument('--body-file', help='Replace entire HTML body from file path')
    parser_edit_html.add_argument('--copy-attachments', action='store_true', help='Copy non-embedded attachments from source')
    parser_edit_html.set_defaults(func=cmd_edit_html)

    # Get OOO status command
    parser_get_ooo = subparsers.add_parser('get-ooo', help='Get Out of Office status')
    parser_get_ooo.set_defaults(func=cmd_get_ooo)

    # Set OOO command
    parser_set_ooo = subparsers.add_parser('set-ooo', help='Enable Out of Office auto-reply')
    parser_set_ooo.set_defaults(func=cmd_set_ooo)

    # Disable OOO command
    parser_disable_ooo = subparsers.add_parser('disable-ooo', help='Disable Out of Office auto-reply')
    parser_disable_ooo.set_defaults(func=cmd_disable_ooo)

    # Get signature command
    parser_get_sig = subparsers.add_parser('get-signature', help='List all signatures or view a specific one')
    parser_get_sig.add_argument('name', nargs='?', default=None, help='Signature name (omit to list all)')
    parser_get_sig.add_argument('--format', choices=['text', 'html'], default='text', help='Output format (default: text)')
    parser_get_sig.set_defaults(func=cmd_get_signature)

    # Update signature command
    parser_update_sig = subparsers.add_parser('update-signature', help='Modify a signature')
    parser_update_sig.add_argument('name', help='Signature name')
    parser_update_sig.add_argument('--text', help='Full plain text replacement (use \\n for line breaks)')
    parser_update_sig.add_argument('--body', help='Full HTML replacement for the signature')
    parser_update_sig.add_argument('--find', help='Text to find (use with --replace)')
    parser_update_sig.add_argument('--replace', help='Replacement text (use with --find)')
    parser_update_sig.add_argument('--after', help='Insert new content after this text (use with --insert)')
    parser_update_sig.add_argument('--insert', help='Text to insert (use \\n for new lines, use with --after)')
    parser_update_sig.set_defaults(func=cmd_update_signature)

    # List calendar command
    parser_calendar = subparsers.add_parser('list-calendar', help='View calendar appointments')
    parser_calendar.add_argument('--user', '--contact', dest='user', help='Check calendar for another user/contact')
    parser_calendar.add_argument('--date', help='Specific date (YYYY-MM-DD)')
    parser_calendar.add_argument('--days', type=int, default=1, help='Days ahead to show (default: today only)')
    parser_calendar.add_argument('--start', help='Start date (YYYY-MM-DD)')
    parser_calendar.add_argument('--end', help='End date (YYYY-MM-DD)')
    parser_calendar.add_argument('--interval', type=int, choices=[15, 30, 60], default=30, help='Slot interval in minutes (default: 30)')
    parser_calendar.add_argument('--working-hours', default='09:00-18:00', help='Working hours range (default: 09:00-18:00, or "all")')
    parser_calendar.add_argument('--compare', action='store_true', help='Compare with your own calendar')
    parser_calendar.add_argument('--json', action='store_true', help='Output JSON')
    parser_calendar.set_defaults(func=cmd_list_calendar)

    # Check availability / free-busy command
    parser_avail = subparsers.add_parser(
        'check-availability',
        aliases=['free-busy', 'freebusy', 'check-calendar'],
        help='Check calendar availability / free-busy status for one or more contacts',
    )
    parser_avail.add_argument('users', nargs='*', default=[], help='One or more email addresses or contact names')
    parser_avail.add_argument('--user', '--contact', dest='user_flag', help='User email or name (alternative to positional)')
    parser_avail.add_argument('--file', help='Path to file containing email addresses (one per line or comma-separated)')
    parser_avail.add_argument('--date', help='Specific date (YYYY-MM-DD)')
    parser_avail.add_argument('--start', help='Start date (YYYY-MM-DD)')
    parser_avail.add_argument('--end', help='End date (YYYY-MM-DD)')
    parser_avail.add_argument('--days', type=int, default=1, help='Number of days to check (default: 1)')
    parser_avail.add_argument('--interval', type=int, choices=[15, 30, 60], default=30, help='Slot interval in minutes (default: 30)')
    parser_avail.add_argument('--working-hours', default='09:00-18:00', help='Working hours range (default: 09:00-18:00, or "all")')
    parser_avail.add_argument('--duration', '--min-duration', dest='duration', type=int, help='Minimum required duration in minutes for common free slots (e.g. 30, 45, 60)')
    parser_avail.add_argument('--compare', action='store_true', help='Compare with your own calendar to find mutual free slots')
    parser_avail.add_argument('--common-only', '--slots-only', dest='common_only', action='store_true', help='Show only the mutual common free slots list')
    parser_avail.add_argument('--json', action='store_true', help='Output structured JSON')
    parser_avail.set_defaults(func=cmd_check_availability)

    # Respond to meeting command
    parser_respond = subparsers.add_parser('respond-meeting', help='Accept/decline/tentative a meeting invite')
    parser_respond.add_argument('email_id', help='Meeting item Entry ID')
    parser_respond.add_argument('--action', required=True, choices=['accept', 'tentative', 'decline', 'propose'])
    parser_respond.add_argument('--start', help='Proposed new start time (YYYY-MM-DD HH:MM, for --action propose)')
    parser_respond.add_argument('--end', help='Proposed new end time (YYYY-MM-DD HH:MM, for --action propose)')
    parser_respond.set_defaults(func=cmd_respond_meeting)

    # Cancel a meeting you organise: sends the cancellation to every invitee.
    # Invitations themselves are created by teams_web_skill (only a browser can
    # mint a Teams join link on this account); cancelling lives here so it uses
    # the item Outlook already holds and the existing review gate.
    parser_cancel = subparsers.add_parser(
        'cancel-invite',
        help='Cancel a meeting you organise, notifying every invitee',
    )
    parser_cancel.add_argument('email_id', help='Appointment Entry ID')
    parser_cancel.add_argument('--subject', help='Resolve by subject when the EntryID is short or unknown')
    parser_cancel.add_argument('--dry-run', dest='dry_run', action='store_true', help='Show the plan, send nothing')
    parser_cancel.set_defaults(func=cmd_cancel_invite)

    # Delete a calendar item without notifying anybody.
    parser_delete_appt = subparsers.add_parser(
        'delete-appointment',
        help='Delete a calendar item without notifying anybody',
    )
    parser_delete_appt.add_argument('email_id', help='Appointment Entry ID')
    parser_delete_appt.add_argument('--subject', help='Resolve by subject when the EntryID is short or unknown')
    parser_delete_appt.add_argument('--force', action='store_true', help='Delete even a live meeting (invitees are not told)')
    parser_delete_appt.add_argument('--dry-run', dest='dry_run', action='store_true', help='Show the plan, delete nothing')
    parser_delete_appt.set_defaults(func=cmd_delete_appointment)

    return parser
