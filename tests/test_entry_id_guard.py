"""Short/truncated EntryIDs must fail fast before reaching MAPI.

An LLM asked to reproduce a ~140-char EntryID can silently drop characters
(observed: a 160-char ID regressed to 90), so _get_email_item rejects short
IDs up front instead of handing a known-bad ID to COM.
"""
import pytest

from backend import email_composition
from backend.cli import common as cli_common


class _FakeNamespace:
    """Every GetItemFromID fails, like a bad or cold COM session."""

    def GetItemFromID(self, *args, **kwargs):
        raise RuntimeError("GetItemFromID failed")


class _FakeSession:
    @property
    def namespace(self):
        return _FakeNamespace()


def test_cli_rejects_short_id():
    with pytest.raises(ValueError, match="140"):
        cli_common._get_email_item(None, "ABC123")


def test_cli_rejects_90_char_id_like_model_hallucination():
    with pytest.raises(ValueError):
        cli_common._get_email_item(None, "0" * 90)


def test_cli_short_id_with_subject_goes_to_fallback_not_guard():
    # --subject is the designed escape hatch: the guard must not fire.
    with pytest.raises(RuntimeError):
        cli_common._get_email_item(_FakeSession(), "ABC123", subject="Some subject")


def test_cli_full_length_id_reaches_outlook():
    with pytest.raises(RuntimeError):
        cli_common._get_email_item(_FakeSession(), "1" * 140)


def test_cli_guard_does_not_touch_session():
    # The guard fires before any COM access, so a None session is fine.
    with pytest.raises(ValueError):
        cli_common._get_email_item(None, "")


def test_library_rejects_short_id():
    with pytest.raises(ValueError, match="140"):
        email_composition._get_email_item(None, "0" * 90)


def test_library_full_length_id_reaches_outlook():
    with pytest.raises(RuntimeError):
        email_composition._get_email_item(_FakeSession(), "1" * 140)
