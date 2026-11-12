"""Folder management and destructive mail commands."""

from __future__ import annotations

import sys

from backend.outlook_session.session_manager import OutlookSessionManager

from .common import _get_email_item


def cmd_create_folder(args):
    """Create a new folder"""
    try:
        from backend.outlook_session.folder_operations import create_folder
        result = create_folder(args.name, args.parent)
        print(result)
        return 0
    except Exception as e:
        print(f"Error: {str(e)}", file=sys.stderr)
        return 1

def cmd_remove_folder(args):
    """Remove a folder"""
    try:
        from backend.outlook_session.folder_operations import remove_folder
        result = remove_folder(args.name)
        print(result)
        return 0
    except Exception as e:
        print(f"Error: {str(e)}", file=sys.stderr)
        return 1

def cmd_move_email(args):
    """Move an email to a folder by email ID"""
    try:
        with OutlookSessionManager() as session:
            email_item = _get_email_item(session, args.email_id)
            
            # Get target folder
            target_folder = session.get_folder(args.folder)
            if not target_folder:
                print(f"Error: Folder '{args.folder}' not found", file=sys.stderr)
                return 1
            
            # Move email
            email_item.Move(target_folder)
            print(f"Successfully moved email to '{args.folder}'")
        
        return 0
    except Exception as e:
        print(f"Error: {str(e)}", file=sys.stderr)
        return 1

def cmd_delete_email(args):
    """Delete an email by ID"""
    try:
        with OutlookSessionManager() as session:
            email_item = _get_email_item(session, args.email_id)
            subject = email_item.Subject
            email_item.Delete()
            print(f"Successfully deleted email: {subject}")

        return 0
    except Exception as e:
        print(f"Error: {str(e)}", file=sys.stderr)
        return 1

def cmd_recall(args):
    """Recall a sent email via Exchange server"""
    try:
        with OutlookSessionManager() as session:
            email_item = _get_email_item(session, args.email_id)

            if not email_item.Sent:
                print("Error: Can only recall sent emails.", file=sys.stderr)
                return 1

            subject = email_item.Subject
            recipients = email_item.To
            print(f"Recalling: {subject}")
            print(f"Recipients: {recipients}")

            email_item.Display()
            inspector = email_item.GetInspector
            inspector.CommandBars.ExecuteMso("RecallThisMessage")
            print(f"\n✅ Recall dialog opened for: {subject}")
            print("Please confirm in the Outlook dialog, then close the email window.")
            print("Note: Recall only works for internal Exchange recipients who haven't read the message.")

        return 0
    except Exception as e:
        print(f"Error: {str(e)}", file=sys.stderr)
        return 1
