"""Stable CLI entrypoint for the Outlook COM skill."""

from pathlib import Path
import sys

SKILL_ROOT = Path(__file__).resolve().parents[1]
if str(SKILL_ROOT) not in sys.path:
    sys.path.insert(0, str(SKILL_ROOT))

from backend.cli.app import main
from backend.cli.runtime import configure_utf8_streams


if __name__ == "__main__":
    configure_utf8_streams()
    raise SystemExit(main())
