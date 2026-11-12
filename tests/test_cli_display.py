from backend.cli.display import _meeting_badge


def test_sent_invite_copy_is_marked_via_message_class():
    # The organizer's Sent Items copy of a meeting request carries no
    # MeetingStatus — MessageClass is the only reliable signal.
    result = _meeting_badge({
        "message_class": "IPM.Schedule.Meeting.Request",
        "meeting_status": "",
    })

    assert result == " 📅 Meeting Invite"


def test_meeting_response_types():
    assert _meeting_badge({"message_class": "IPM.Schedule.Meeting.Resp.Pos"}) == " ✅ Accepted"
    assert _meeting_badge({"message_class": "IPM.Schedule.Meeting.Resp.Neg"}) == " ❌ Declined"
    assert _meeting_badge({"message_class": "IPM.Schedule.Meeting.Resp.Tent"}) == " 📅 Tentative"
    assert _meeting_badge({"message_class": "IPM.Schedule.Meeting.Canceled"}) == " ❌ Canceled"
    assert _meeting_badge({"message_class": "IPM.Schedule.Meeting.Msg"}) == " 📅 Meeting Update"


def test_meeting_status_fallback_when_message_class_is_plain():
    result = _meeting_badge({"message_class": "IPM.Note", "meeting_status": "meeting_request"})

    assert result == " 📅 Meeting Invite"


def test_plain_email_has_no_badge():
    result = _meeting_badge({
        "message_class": "IPM.Note",
        "meeting_status": "",
        "subject": "Q4 plan review",
        "sender": "Hong Yang",
    })

    assert result == ""


def test_event_heuristic_still_applies():
    result = _meeting_badge({"subject": "You are invited: AI webinar"})

    assert result == " 📅 Event"
