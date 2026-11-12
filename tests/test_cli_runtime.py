import argparse

import pytest

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


def test_preflight_blocks_low_integrity_process(monkeypatch):
    monkeypatch.delenv("OUTLOOK_SKILL_IGNORE_INTEGRITY", raising=False)
    monkeypatch.setattr(runtime, "is_low_integrity", lambda: True)
    error = runtime._outlook_preflight_error()
    assert error is not None
    assert "Low integrity" in error


def test_preflight_allows_medium_integrity_process(monkeypatch):
    monkeypatch.delenv("OUTLOOK_SKILL_IGNORE_INTEGRITY", raising=False)
    monkeypatch.setattr(runtime, "is_low_integrity", lambda: False)
    assert runtime._outlook_preflight_error() is None


def test_preflight_can_be_overridden_for_diagnostics(monkeypatch):
    monkeypatch.setenv("OUTLOOK_SKILL_IGNORE_INTEGRITY", "1")
    monkeypatch.setattr(runtime, "is_low_integrity", lambda: True)
    assert runtime._outlook_preflight_error() is None


def test_preflight_never_raises_when_probe_fails(monkeypatch):
    monkeypatch.delenv("OUTLOOK_SKILL_IGNORE_INTEGRITY", raising=False)

    def boom():
        raise OSError("probe exploded")

    monkeypatch.setattr(runtime, "is_low_integrity", boom)
    assert runtime._outlook_preflight_error() is None


def test_isolated_run_fails_fast_on_low_integrity(monkeypatch, capsys):
    monkeypatch.delenv(runtime.OUTLOOK_SKILL_CHILD_ENV, raising=False)
    monkeypatch.delenv("OUTLOOK_SKILL_IGNORE_INTEGRITY", raising=False)
    monkeypatch.setattr(runtime, "is_low_integrity", lambda: True)

    def explode(*args, **kwargs):
        raise AssertionError("must not spawn a child process at Low integrity")

    monkeypatch.setattr(runtime.subprocess, "run", explode)
    args = argparse.Namespace(command="find", timeout=None)

    assert runtime._run_outlook_command_isolated(args) == 125
    assert "cannot reach Outlook COM" in capsys.readouterr().err
