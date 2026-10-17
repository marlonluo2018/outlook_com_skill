from datetime import datetime

from backend.cli.search_commands import _serialize_recent_email
from backend.email_search.parallel_extractor import extract_emails_lightweight


def test_recent_json_contract_includes_conversation_id():
    result = _serialize_recent_email({
        "id": "ENTRY-1",
        "conversation_id": "CONV-1",
        "folder_name": "Inbox",
    })

    assert result["entry_id"] == "ENTRY-1"
    assert result["conversation_id"] == "CONV-1"
    assert result["folder"] == "Inbox"


def test_recent_json_contract_defaults_missing_conversation_id():
    result = _serialize_recent_email({"entry_id": "ENTRY-2"})

    assert result["conversation_id"] == ""


class _FakeMail:
    EntryID = "ENTRY-3"
    ConversationID = "CONV-3"
    Subject = "Subject"
    SenderName = "Sender"
    SenderEmailAddress = "sender@example.com"
    ReceivedTime = datetime(2026, 9, 22, 9, 30)
    MessageClass = "IPM.Note"
    MeetingStatus = 0
    Recipients = None
    Attachments = None
    Body = "Preview"
    UnRead = False


def test_lightweight_extractor_preserves_conversation_id():
    result = extract_emails_lightweight([_FakeMail()])

    assert result[0]["entry_id"] == "ENTRY-3"
    assert result[0]["conversation_id"] == "CONV-3"
