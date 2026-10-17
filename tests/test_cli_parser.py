from backend.cli.parser import build_parser
from backend.config import search_config


EXPECTED_COMMANDS = {
    "list-folders",
    "find-recent",
    "find",
    "get-email",
    "read",
    "reply",
    "compose",
    "forward",
    "redirect",
    "batch-forward",
    "create-folder",
    "remove-folder",
    "move-email",
    "delete-email",
    "recall",
    "lookup-contact",
    "find-thread",
    "find-related",
    "download-attachment",
    "get-html",
    "send-draft",
    "edit-html",
    "get-ooo",
    "set-ooo",
    "disable-ooo",
    "get-signature",
    "update-signature",
    "list-calendar",
    "respond-meeting",
    "check-availability",
    "free-busy",
    "freebusy",
    "check-calendar",
}


def test_parser_exposes_all_commands():
    parser = build_parser()
    subparsers = parser._subparsers._group_actions[0]
    assert set(subparsers.choices) == EXPECTED_COMMANDS


def test_read_alias_uses_get_email_handler():
    parser = build_parser()
    primary = parser.parse_args(["get-email", "abc"])
    alias = parser.parse_args(["read", "abc"])

    assert alias.command == "read"
    assert alias.email_ids == ["abc"]
    assert alias.func is primary.func


def test_find_defaults_match_search_config():
    args = build_parser().parse_args(["find", "--subject", "voucher"])

    assert args.subject == "voucher"
    assert args.days == search_config.DIRECT_FIND_DEFAULT_DAYS
    assert args.limit is None
    assert args.match_all is True


def test_find_related_limit_maps_to_max_results():
    args = build_parser().parse_args(["find-related", "entry-id", "--limit", "7"])

    assert args.max_results == 7


def test_reply_repeated_cc_is_preserved_for_normalization():
    args = build_parser().parse_args(
        ["reply", "entry-id", "hello", "--cc", "a@example.com", "--cc", "b@example.com"]
    )

    assert args.cc == ["a@example.com", "b@example.com"]


def test_check_availability_parser_options():
    parser = build_parser()
    args = parser.parse_args([
        "check-availability",
        "user1@ibm.com",
        "user2@ibm.com",
        "--date", "2026-09-28",
        "--interval", "15",
        "--working-hours", "08:30-17:30",
        "--duration", "60",
        "--common-only",
        "--compare",
        "--json",
    ])

    assert args.command == "check-availability"
    assert args.users == ["user1@ibm.com", "user2@ibm.com"]
    assert args.date == "2026-09-28"
    assert args.interval == 15
    assert args.working_hours == "08:30-17:30"
    assert args.duration == 60
    assert args.common_only is True
    assert args.compare is True
    assert args.json is True


def test_check_availability_aliases():
    parser = build_parser()
    for alias in ["free-busy", "freebusy", "check-calendar"]:
        args = parser.parse_args([alias, "user@ibm.com"])
        assert args.command == alias
        assert args.users == ["user@ibm.com"]
        assert args.func is not None


def test_list_calendar_user_argument():
    parser = build_parser()
    args = parser.parse_args(["list-calendar", "--user", "user@ibm.com", "--compare"])
    assert args.user == "user@ibm.com"
    assert args.compare is True
