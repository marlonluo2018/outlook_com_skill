import argparse

from backend.cli import runtime


def test_timeout_uses_command_specific_value():
    args = argparse.Namespace(command="reply", timeout=None)
    assert runtime._timeout_for_command(args) == 90


def test_timeout_uses_default_for_unknown_command():
    args = argparse.Namespace(command="unknown", timeout=None)
    assert runtime._timeout_for_command(args) == runtime.OUTLOOK_DEFAULT_TIMEOUT_SECONDS


def test_timeout_override_is_clamped_to_one_second():
    args = argparse.Namespace(command="lookup-contact", timeout=0)
    assert runtime._timeout_for_command(args) == 1


def test_isolation_skips_child_process(monkeypatch):
    args = argparse.Namespace(command="find")
    monkeypatch.setenv(runtime.OUTLOOK_SKILL_CHILD_ENV, "1")
    assert runtime._should_isolate_outlook_command(args) is False


def test_lock_is_a_context_manager_when_windows_locking_is_unavailable(monkeypatch):
    monkeypatch.setattr(runtime, "msvcrt", None)
    with runtime._outlook_cli_lock():
        pass
