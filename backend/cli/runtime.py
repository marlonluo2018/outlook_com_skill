"""CLI runtime safeguards for Outlook COM commands."""

from __future__ import annotations

import contextlib
import os
from pathlib import Path
import subprocess
import sys
import tempfile
import time

try:
    import msvcrt
except ImportError:  # pragma: no cover - this Outlook skill is Windows-only
    msvcrt = None

try:
    from backend.integrity import is_low_integrity
except ImportError:  # pragma: no cover - direct package-relative import
    from ..integrity import is_low_integrity


OUTLOOK_SKILL_CHILD_ENV = "OUTLOOK_SKILL_CHILD"
OUTLOOK_LOCK_TIMEOUT_SECONDS = int(os.environ.get("OUTLOOK_SKILL_LOCK_TIMEOUT", "30"))
OUTLOOK_DEFAULT_TIMEOUT_SECONDS = int(os.environ.get("OUTLOOK_SKILL_TIMEOUT", "60"))
OUTLOOK_COMMAND_TIMEOUTS = {
    "lookup-contact": 20,
    "compose": 90,
    "reply": 90,
    "forward": 90,
    "redirect": 90,
    "send-draft": 90,
    "batch-forward": 300,
    "batch-compose": 300,
    "recall": 90,
    "respond-meeting": 90,
    "download-attachment": 120,
}


def configure_utf8_streams() -> None:
    """Use UTF-8 for CLI streams when the host supports reconfiguration."""
    for stream_name in ("stdin", "stdout", "stderr"):
        stream = getattr(sys, stream_name, None)
        if stream and hasattr(stream, "reconfigure") and stream.encoding != "utf-8":
            stream.reconfigure(encoding="utf-8")


class OutlookCliLockTimeout(TimeoutError):
    """Raised when another Outlook COM command holds the process-wide lock."""

@contextlib.contextmanager
def _outlook_cli_lock(timeout_seconds: int = OUTLOOK_LOCK_TIMEOUT_SECONDS):
    """Serialize Outlook COM access across CLI processes.

    Outlook COM is single-user desktop automation and becomes unreliable when
    multiple Python processes drive it at once. A byte-range lock gives us an
    OS-released mutex even if a process is killed after a timeout.
    """
    if msvcrt is None:
        yield
        return

    lock_path = os.path.join(tempfile.gettempdir(), "brainclaw_outlook_com.lock")
    with open(lock_path, "a+b") as lock_file:
        try:
            lock_file.seek(0)
            if not lock_file.read(1):
                lock_file.write(b"0")
                lock_file.flush()
        except OSError:
            # Another process holds the lock on byte 0; proceed to locking loop below.
            pass

        deadline = time.monotonic() + max(timeout_seconds, 1)
        while True:
            try:
                lock_file.seek(0)
                msvcrt.locking(lock_file.fileno(), msvcrt.LK_NBLCK, 1)
                break
            except OSError:
                if time.monotonic() >= deadline:
                    raise OutlookCliLockTimeout(
                        f"Timed out waiting {timeout_seconds}s for another Outlook command to finish."
                    )
                time.sleep(0.25)

        try:
            yield
        finally:
            lock_file.seek(0)
            try:
                msvcrt.locking(lock_file.fileno(), msvcrt.LK_UNLCK, 1)
            except OSError:
                pass

def _should_isolate_outlook_command(args) -> bool:
    """Return True when the parsed command should run in a timeout child process."""
    return bool(getattr(args, "command", None)) and os.environ.get(OUTLOOK_SKILL_CHILD_ENV) != "1"

def _timeout_for_command(args) -> int:
    requested = getattr(args, "timeout", None)
    if requested is not None:
        return max(int(requested), 1)
    return OUTLOOK_COMMAND_TIMEOUTS.get(args.command, OUTLOOK_DEFAULT_TIMEOUT_SECONDS)

LOW_INTEGRITY_HINT = (
    "This process is running at Low integrity (a sandboxed/non-elevated tool context). "
    "Windows blocks cross-integrity Outlook COM calls, so Outlook instead tries to start a "
    "second instance and shows a modal dialog, which blocks the command until it times out "
    "and leaves that dialog on the desktop.\n"
    "Run this command from the interactive desktop session (Medium integrity) instead, then "
    "dismiss any 'Only one version of Outlook can run at a time' dialog."
)

def _outlook_preflight_error() -> str | None:
    """Return an actionable error when this process cannot drive Outlook COM at all."""
    if os.environ.get("OUTLOOK_SKILL_IGNORE_INTEGRITY") == "1":
        return None
    try:
        if is_low_integrity():
            return LOW_INTEGRITY_HINT
    except Exception:  # pragma: no cover - never block on the guard itself
        return None
    return None

def _run_outlook_command_isolated(args) -> int:
    """Run the command in a child process so blocking COM calls can be terminated."""
    timeout_seconds = _timeout_for_command(args)
    env = os.environ.copy()
    env[OUTLOOK_SKILL_CHILD_ENV] = "1"

    preflight_error = _outlook_preflight_error()
    if preflight_error:
        print(f"Error: cannot reach Outlook COM from this process.", file=sys.stderr)
        print(preflight_error, file=sys.stderr)
        return 125

    entrypoint = Path(__file__).resolve().parents[2] / "scripts" / "outlook_skill.py"
    command = [sys.executable, str(entrypoint), *sys.argv[1:]]
    try:
        with _outlook_cli_lock():
            completed = subprocess.run(
                command,
                env=env,
                timeout=timeout_seconds,
                text=True,
                encoding="utf-8",
                errors="replace",
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
            )
    except OutlookCliLockTimeout as e:
        print(f"Error: {e}", file=sys.stderr)
        return 124
    except subprocess.TimeoutExpired as e:
        stdout = e.stdout or ""
        stderr = e.stderr or ""
        if stdout:
            print(stdout, end="")
        if stderr:
            print(stderr, end="", file=sys.stderr)
        print(
            f"Error: Outlook command '{args.command}' timed out after {timeout_seconds}s and was stopped.",
            file=sys.stderr,
        )
        print(
            "Outlook COM may be busy, waiting on a hidden dialog, or unable to access the current desktop session.",
            file=sys.stderr,
        )
        return 124

    if completed.stdout:
        print(completed.stdout, end="")
    if completed.stderr:
        print(completed.stderr, end="", file=sys.stderr)
    return completed.returncode
