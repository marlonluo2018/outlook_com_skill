"""Application entrypoint for the Outlook COM skill CLI."""

from __future__ import annotations

import sys
from collections.abc import Sequence
from pathlib import Path
from typing import Optional

from backend.config import HostPathConfig

from .parser import build_parser
from .runtime import _run_outlook_command_isolated, _should_isolate_outlook_command

# Temporary email bodies live here and are removed after a successful send.
OUTBOX_DIR = HostPathConfig.outbox_dir()
SEND_COMMANDS = {"reply", "compose", "forward", "redirect", "batch-forward", "batch-compose"}


def _cleanup_outbox_body(args) -> None:
    """Delete a sent body file, but only when it lives inside OUTBOX_DIR."""
    body_file = getattr(args, "body_file", None)
    if not body_file or args.command not in SEND_COMMANDS:
        return
    path = Path(body_file).resolve()
    if OUTBOX_DIR.resolve() not in path.parents:
        return
    try:
        path.unlink(missing_ok=True)
        print(f"[outbox] removed sent body file: {path.name}", file=sys.stderr)
    except OSError as exc:
        print(f"Warning: could not remove body file {path}: {exc}", file=sys.stderr)


def main(argv: Optional[Sequence[str]] = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)

    if not args.command:
        parser.print_help()
        return 1

    if _should_isolate_outlook_command(args):
        return _run_outlook_command_isolated(args)

    result = args.func(args)
    if result == 0:
        _cleanup_outbox_body(args)
    return result
