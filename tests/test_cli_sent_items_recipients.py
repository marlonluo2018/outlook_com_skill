"""Tests for Sent Items recipient copying in reply commands.

Covers _copy_sent_items_recipients, which backs both reply --only and
reply-all when the source email lives in Sent Items. Fake objects mimic the
small Outlook COM surface the helper touches (Recipients.Count/Item/Add,
Recipient.Address/Type); _get_smtp_address falls back to .Address when the
AddressEntry lookup fails on the fakes.
"""

from backend.cli.common import _copy_sent_items_recipients


class _FakeRecipient:
    def __init__(self, address, recip_type):
        self.Address = address
        self.Type = recip_type
        self.Name = address


class _FakeRecipients:
    def __init__(self, items=None):
        self._items = list(items or [])

    @property
    def Count(self):
        return len(self._items)

    def Item(self, i):
        return self._items[i - 1]  # Outlook COM is 1-based

    def Add(self, smtp):
        recip = _FakeRecipient(smtp, 1)
        self._items.append(recip)
        return recip


class _FakeEmail:
    def __init__(self, recipients):
        self.Recipients = _FakeRecipients(recipients)


class _FakeMail:
    """Reply item surface used by the helper (a .Recipients collection)."""

    def __init__(self):
        self.Recipients = _FakeRecipients()


CURRENT_USER = "me@example.com"


def _sent_email():
    # A sent item: To a business contact, Cc an LDM, self in To (e.g. via a
    # mailing list expansion), plus another To contact.
    return _FakeEmail([
        _FakeRecipient("to.one@example.com", 1),   # To
        _FakeRecipient("cc.one@example.com", 2),     # Cc
        _FakeRecipient(CURRENT_USER, 1),            # self in To
        _FakeRecipient("to.two@example.com", 1),    # To
    ])


def test_to_only_copies_to_recipients_and_skips_cc_and_self():
    reply = _FakeMail()

    _copy_sent_items_recipients(_sent_email(), reply, CURRENT_USER, to_only=True)

    assert [r.Address for r in reply.Recipients._items] == ["to.one@example.com", "to.two@example.com"]
    assert all(r.Type == 1 for r in reply.Recipients._items)


def test_default_copies_to_and_cc_and_skips_self():
    reply = _FakeMail()

    _copy_sent_items_recipients(_sent_email(), reply, CURRENT_USER)

    assert [r.Address for r in reply.Recipients._items] == [
        "to.one@example.com",
        "cc.one@example.com",
        "to.two@example.com",
    ]
    # Original To/Cc types are preserved on the new item.
    assert [r.Type for r in reply.Recipients._items] == [1, 2, 1]


def test_self_comparison_is_case_insensitive():
    reply = _FakeMail()
    email = _FakeEmail([
        _FakeRecipient("TO.ONE@EXAMPLE.COM", 1),
        _FakeRecipient("ME@EXAMPLE.COM", 1),
    ])

    _copy_sent_items_recipients(email, reply, CURRENT_USER, to_only=True)

    assert [r.Address for r in reply.Recipients._items] == ["TO.ONE@EXAMPLE.COM"]


def test_recipient_missing_type_defaults_to_to():
    reply = _FakeMail()
    recip = _FakeRecipient("to.one@example.com", 1)
    del recip.Type  # attribute absent: getattr default treats it as To
    email = _FakeEmail([recip])

    _copy_sent_items_recipients(email, reply, CURRENT_USER, to_only=True)

    assert [r.Address for r in reply.Recipients._items] == ["to.one@example.com"]
