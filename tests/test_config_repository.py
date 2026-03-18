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
    assert (tmp_path / "config.json.bak").exists()


def test_save_and_reload_preserves_known_keys(tmp_path: Path) -> None:
    config_file = tmp_path / "config.json"
    repo = ConfigRepository(config_file)
    config = AppConfig(theme="light", format="DOCX", folder_path="C:/work")

    repo.save(config)
    reloaded = repo.load()

    assert reloaded.theme == "light"
    assert reloaded.format == "DOCX"
    assert reloaded.folder_path == "C:/work"
