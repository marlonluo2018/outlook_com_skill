from pathlib import Path
import subprocess
import sys


def test_help_entrypoint_does_not_open_outlook():
    root = Path(__file__).resolve().parents[1]
    completed = subprocess.run(
        [sys.executable, str(root / "scripts" / "outlook_skill.py"), "--help"],
        cwd=root,
        capture_output=True,
        text=True,
        encoding="utf-8",
        timeout=20,
        check=False,
    )

    assert completed.returncode == 0
    assert "find-recent" in completed.stdout
    assert "respond-meeting" in completed.stdout
    assert completed.stderr == ""
