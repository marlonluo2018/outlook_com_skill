from datetime import datetime, time
from backend.cli.calendar_commands import (
    _parse_working_hours,
    _get_date_list,
    _group_contiguous_blocks,
    _compute_mutual_free_blocks,
)


class DummyArgs:
    def __init__(self, **kwargs):
        for k, v in kwargs.items():
            setattr(self, k, v)


def test_parse_working_hours():
    assert _parse_working_hours("09:00-18:00") == (time(9, 0), time(18, 0))
    assert _parse_working_hours("all") is None
    assert _parse_working_hours("24h") is None
    assert _parse_working_hours("invalid") == (time(9, 0), time(18, 0))


def test_get_date_list_with_date():
    args = DummyArgs(date="2026-09-28")
    dates = _get_date_list(args)
    assert len(dates) == 1
    assert dates[0].strftime("%Y-%m-%d") == "2026-09-28"


def test_get_date_list_with_start_end():
    args = DummyArgs(start="2026-09-28", end="2026-09-30")
    dates = _get_date_list(args)
    assert len(dates) == 3
    assert [d.strftime("%Y-%m-%d") for d in dates] == ["2026-09-28", "2026-09-29", "2026-09-30"]


def test_group_contiguous_blocks():
    slots = [
        {"start": "09:00", "end": "09:30", "start_time": time(9, 0), "end_time": time(9, 30), "status": "Free"},
        {"start": "09:30", "end": "10:00", "start_time": time(9, 30), "end_time": time(10, 0), "status": "Free"},
        {"start": "10:00", "end": "10:30", "start_time": time(10, 0), "end_time": time(10, 30), "status": "Busy"},
        {"start": "10:30", "end": "11:00", "start_time": time(10, 30), "end_time": time(11, 0), "status": "Busy"},
        {"start": "11:00", "end": "12:00", "start_time": time(11, 0), "end_time": time(12, 0), "status": "Free"},
    ]
    blocks = _group_contiguous_blocks(slots)
    assert len(blocks) == 3
    assert blocks[0] == {"start": "09:00", "end": "10:00", "status": "Free", "duration_minutes": 60, "slot_count": 2}
    assert blocks[1] == {"start": "10:00", "end": "11:00", "status": "Busy", "duration_minutes": 60, "slot_count": 2}
    assert blocks[2] == {"start": "11:00", "end": "12:00", "status": "Free", "duration_minutes": 60, "slot_count": 1}


def test_compute_mutual_free_blocks():
    # User 1: Free 09:00-11:00, Busy 11:00-12:00
    user1_slots = [
        {"start": "09:00", "end": "10:00", "start_time": time(9, 0), "end_time": time(10, 0), "status": "Free"},
        {"start": "10:00", "end": "11:00", "start_time": time(10, 0), "end_time": time(11, 0), "status": "Free"},
        {"start": "11:00", "end": "12:00", "start_time": time(11, 0), "end_time": time(12, 0), "status": "Busy"},
    ]
    # User 2: Busy 09:00-10:00, Free 10:00-12:00
    user2_slots = [
        {"start": "09:00", "end": "10:00", "start_time": time(9, 0), "end_time": time(10, 0), "status": "Busy"},
        {"start": "10:00", "end": "11:00", "start_time": time(10, 0), "end_time": time(11, 0), "status": "Free"},
        {"start": "11:00", "end": "12:00", "start_time": time(11, 0), "end_time": time(12, 0), "status": "Free"},
    ]

    mutual = _compute_mutual_free_blocks([user1_slots, user2_slots], interval=60)
    assert len(mutual) == 1
    assert mutual[0]["start"] == "10:00"
    assert mutual[0]["end"] == "11:00"
    assert mutual[0]["duration_minutes"] == 60
