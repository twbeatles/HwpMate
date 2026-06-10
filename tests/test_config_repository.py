from __future__ import annotations

from pathlib import Path

from hwpmate.config_repository import ConfigRepository
from hwpmate.models import AppConfig


def test_load_merges_defaults(tmp_path: Path) -> None:
    config_file = tmp_path / "config.json"
    config_file.write_text('{"theme": "light", "last_folder": "C:/docs"}', encoding="utf-8")

    repo = ConfigRepository(config_file)
    config = repo.load()

    assert isinstance(config, AppConfig)
    assert config.theme == "light"
    assert config.mode == "folder"
    assert config.last_folder == "C:/docs"


def test_load_backs_up_invalid_json(tmp_path: Path) -> None:
    config_file = tmp_path / "config.json"
    config_file.write_text("{ invalid", encoding="utf-8")

    repo = ConfigRepository(config_file)
    config = repo.load()

    assert config.theme == "dark"
    backups = list(tmp_path.glob("config_*.json.bak"))
    assert len(backups) == 1
    assert not config_file.exists()


def test_save_and_reload_preserves_known_keys(tmp_path: Path) -> None:
    config_file = tmp_path / "config.json"
    repo = ConfigRepository(config_file)
    config = AppConfig(theme="light", format="DOCX", folder_path="C:/work")

    assert repo.save(config) is True
    reloaded = repo.load()

    assert reloaded.theme == "light"
    assert reloaded.format == "DOCX"
    assert reloaded.folder_path == "C:/work"
    assert reloaded.backup_enabled is True
    assert reloaded.retry_count == 1


def test_save_replaces_config_atomically_without_leaving_temp_files(tmp_path: Path) -> None:
    config_file = tmp_path / "config.json"
    repo = ConfigRepository(config_file)

    assert repo.save(AppConfig(theme="light", backup_enabled=False, retry_count=2)) is True

    assert config_file.exists()
    assert not list(tmp_path.glob(".config.json.*.tmp"))
    reloaded = repo.load()
    assert reloaded.backup_enabled is False
    assert reloaded.retry_count == 2


def test_load_recovers_type_mismatches(tmp_path: Path) -> None:
    config_file = tmp_path / "config.json"
    config_file.write_text(
        """
        {
          "config_version": "bad",
          "theme": 123,
          "mode": "files",
          "format": "NOPE",
          "include_sub": "false",
          "same_location": "false",
          "overwrite": "yes",
          "backup_enabled": "off",
          "retry_count": "abc",
          "folder_path": 42
        }
        """,
        encoding="utf-8",
    )

    repo = ConfigRepository(config_file)
    config = repo.load()

    assert config.theme == "dark"
    assert config.mode == "files"
    assert config.format == "PDF"
    assert config.include_sub is False
    assert config.same_location is False
    assert config.overwrite is True
    assert config.backup_enabled is False
    assert config.retry_count == 1
    assert config.folder_path == ""


def test_save_reports_failure_when_parent_is_not_directory(tmp_path: Path) -> None:
    blocked_parent = tmp_path / "blocked"
    blocked_parent.write_text("not a directory", encoding="utf-8")
    repo = ConfigRepository(blocked_parent / "config.json")

    assert repo.save(AppConfig(theme="light")) is False
    assert blocked_parent.read_text(encoding="utf-8") == "not a directory"
