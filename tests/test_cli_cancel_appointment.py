"""Cancel/delete behaviour, exercised with fake COM objects.

The real commands talk to Outlook through COM, so these tests patch the session
manager and item resolver and assert on what the command would do: send a
cancellation, refuse a live meeting, or delete silently. Nothing here touches a
mailbox.
"""

import sys
from contextlib import contextmanager
from types import SimpleNamespace
from unittest import mock

import pytest

sys.path.insert(0, ".")

from backend.cli import calendar_commands as cal


class FakeRecipients:
    def __init__(self, addresses):
        self.Count = len(addresses)
        self._items = [SimpleNamespace(Address=a) for a in addresses]

    def Item(self, index):
        return self._items[index - 1]


class FakeItem:
    def __init__(self, status=1, addresses=("someone@example.com",)):
        self.Subject = "Test meeting"
        self.MeetingStatus = status
        self.Recipients = FakeRecipients(list(addresses))
        self.Location = "Microsoft Teams Meeting"
        self.Organizer = "me@example.com"
        self.Start = None
        self.End = None
        self.sent = 0
        self.deleted = 0

    def Send(self):
        self.sent += 1

    def Delete(self):
        self.deleted += 1


@contextmanager
def patched(item):
    """Patch the COM session and the item resolver; nothing touches a mailbox."""
    session = mock.MagicMock()
    session.__enter__ = mock.Mock(return_value=session)
    session.__exit__ = mock.Mock(return_value=False)
    manager = mock.Mock(return_value=session)
    with mock.patch.object(cal, "OutlookSessionManager", manager), \
         mock.patch.object(cal, "_get_email_item", mock.Mock(return_value=item)):
        yield


def _args(**kwargs):
    base = {"email_id": "x" * 160, "subject": None, "dry_run": False, "force": False}
    base.update(kwargs)
    return SimpleNamespace(**base)


def test_cancel_sends_and_marks_cancelled():
    item = FakeItem(status=1)
    with patched(item):
        assert cal.cmd_cancel_invite(_args()) == 0
    assert item.sent == 1
    assert item.MeetingStatus == cal._OL_MEETING_CANCELED
    # the notice goes out and the organiser copy is removed in the same command
    assert item.deleted == 1


def test_cancel_refuses_a_non_meeting():
    item = FakeItem(status=0)
    with patched(item):
        assert cal.cmd_cancel_invite(_args()) == 1
    assert item.sent == 0


def test_cancel_dry_run_sends_nothing():
    item = FakeItem(status=1)
    with patched(item):
        assert cal.cmd_cancel_invite(_args(dry_run=True)) == 0
    assert item.sent == 0
    assert item.MeetingStatus == 1


def test_delete_refuses_a_live_meeting_without_force():
    item = FakeItem(status=1)
    with patched(item):
        assert cal.cmd_delete_appointment(_args()) == 1
    assert item.deleted == 0


def test_delete_with_force_removes_it():
    item = FakeItem(status=1)
    with patched(item):
        assert cal.cmd_delete_appointment(_args(force=True)) == 0
    assert item.deleted == 1


def test_delete_of_a_plain_event_needs_no_force():
    item = FakeItem(status=0)
    with patched(item):
        assert cal.cmd_delete_appointment(_args()) == 0
    assert item.deleted == 1


def test_summary_reports_invitees():
    item = FakeItem(status=1, addresses=("a@example.com", "b@example.com"))
    summary = cal._appointment_summary(item)
    assert summary["recipients"] == ["a@example.com", "b@example.com"]
    assert summary["meeting_status"] == 1
