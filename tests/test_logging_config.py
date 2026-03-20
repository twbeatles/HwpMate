from __future__ import annotations

from pathlib import Path

from hwpmate import logging_config


def test_resolve_log_file_falls_back_when_primary_path_is_file(tmp_path: Path) -> None:
    blocked_dir = tmp_path / "blocked-logs"
    blocked_dir.write_text("not a directory", encoding="utf-8")
    fallback_dir = tmp_path / "fallback" / "logs"

    log_file, error = logging_config._resolve_log_file([blocked_dir, fallback_dir])

    assert error is None
    assert log_file == fallback_dir / logging_config.LOG_FILE_NAME
    assert log_file is not None
    assert log_file.exists()


def test_resolve_log_file_reports_failure_when_all_candidates_are_blocked(tmp_path: Path) -> None:
    blocked_dir_1 = tmp_path / "blocked-1"
    blocked_dir_2 = tmp_path / "blocked-2"
    blocked_dir_1.write_text("not a directory", encoding="utf-8")
    blocked_dir_2.write_text("not a directory", encoding="utf-8")

    log_file, error = logging_config._resolve_log_file([blocked_dir_1, blocked_dir_2])

    assert log_file is None
    assert error is not None
    assert "blocked-1" in error
    assert "blocked-2" in error
