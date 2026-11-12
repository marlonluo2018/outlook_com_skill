"""HostPathConfig: env overrides, host-layout defaults, standalone fallback,
and the find-cache provenance shown on --pick/--last.
"""
import json
import os
import time
from pathlib import Path
from unittest import mock

from backend import config as skill_config
from backend.cli import common as cli_common
from backend.cli import search_commands


def test_env_overrides_each_path(tmp_path):
    with mock.patch.dict(
        os.environ,
        {
            "OUTLOOK_SKILL_TASKS_DIR": str(tmp_path),
            "OUTLOOK_SKILL_CACHE_DIR": str(tmp_path),
            "OUTLOOK_SKILL_OUTBOX_DIR": str(tmp_path),
            "OUTLOOK_SKILL_DOWNLOADS_DIR": str(tmp_path),
        },
    ):
        assert str(skill_config.HostPathConfig.tasks_dir()) == str(tmp_path)
        assert str(skill_config.HostPathConfig.cache_dir()) == str(tmp_path)
        assert str(skill_config.HostPathConfig.outbox_dir()) == str(tmp_path)
        assert str(skill_config.HostPathConfig.downloads_dir()) == str(tmp_path)


def test_host_detected_defaults_follow_host_layout(tmp_path):
    with mock.patch.object(skill_config, "_host_root", return_value=tmp_path):
        cfg = skill_config.HostPathConfig
        assert cfg.tasks_dir() == tmp_path / "assistant_brain" / "tasks"
        assert cfg.cache_dir() == tmp_path / "assistant_brain" / "sync_results"
        assert cfg.outbox_dir() == tmp_path / "outbox"
        assert cfg.downloads_dir() == tmp_path / "downloads"
        assert cfg.inline_images_dir() == tmp_path / "downloads" / "outlook_inline"


def test_standalone_fallback_uses_outlook_skill_dir():
    standalone = skill_config._STANDALONE_ROOT
    with mock.patch.object(skill_config, "_host_root", return_value=None):
        cfg = skill_config.HostPathConfig
        assert cfg.cache_dir() == standalone / "cache"
        assert cfg.tasks_dir() == standalone / "tasks"
        assert cfg.outbox_dir() == standalone / "outbox"
        assert cfg.downloads_dir() == standalone / "downloads"


def test_task_ref_resolves_from_env_tasks_dir(tmp_path):
    entry_id = "A" * 140
    task_file = tmp_path / "T999-demo.md"
    task_file.write_text("- [email-in] demo <!-- email:%s -->\n" % entry_id, encoding="utf-8")
    with mock.patch.dict(os.environ, {"OUTLOOK_SKILL_TASKS_DIR": str(tmp_path)}):
        resolved, _subject, note = search_commands._resolve_reference(None, None, "T999")
    assert resolved == entry_id
    assert "T999-demo.md" in note


def test_pick_note_carries_source_command(tmp_path):
    emails = [
        {
            "id": "B" * 140,
            "subject": "Hello",
            "sender": "A",
            "received_time": "2026-10-28 10:00:00",
            "folder": "Inbox",
        }
    ]
    with mock.patch.dict(os.environ, {"OUTLOOK_SKILL_CACHE_DIR": str(tmp_path)}):
        cli_common.cache_search_results(emails, context="find --from 'fei' --days 7")
        resolved, _subject, note = search_commands._resolve_reference(1, False, None)
    assert resolved == "B" * 140
    assert "#1/1" in note
    assert "find --from 'fei' --days 7" in note
    assert "⚠" not in note  # fresh cache: no staleness warning


def test_stale_cache_warns(tmp_path):
    emails = [
        {
            "id": "C" * 140,
            "subject": "Old",
            "sender": "A",
            "received_time": "2026-10-28 10:00:00",
            "folder": "Inbox",
        }
    ]
    with mock.patch.dict(os.environ, {"OUTLOOK_SKILL_CACHE_DIR": str(tmp_path)}):
        cli_common.cache_search_results(emails, context="find-recent --days 1")
        cache_path = Path(cli_common._search_cache_path())
        data = json.loads(cache_path.read_text(encoding="utf-8"))
        data["saved_at"] = time.time() - 7200  # 2h old
        cache_path.write_text(json.dumps(data), encoding="utf-8")
        _resolved, _subject, note = search_commands._resolve_reference(None, True, None)
    assert "⚠" in note
    assert "find-recent --days 1" in note
