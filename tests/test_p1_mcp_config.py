"""P1-B — MCP 설정 로딩(config) 단위 테스트.

hwpmate/mcp/config.py 커버리지 공백을 메운다.
fail-closed 기본값, TOML 파일([mcp]/[paths]) 병합, 환경 변수 재정의,
알 수 없는 키 무시, job 디렉터리 결정을 검증한다.
"""

from __future__ import annotations

import os
from pathlib import Path

from hwpmate.mcp.config import McpConfig, default_config_path, default_job_dir, load_config


def _write_toml(path: Path, body: str) -> Path:
    path.write_text(body, encoding="utf-8")
    return path


def test_defaults_are_fail_closed() -> None:
    cfg = McpConfig()
    assert cfg.input_roots == [] and cfg.output_roots == []
    assert cfg.require_preview is True
    assert cfg.require_confirmation is True
    assert cfg.allow_overwrite is False
    assert cfg.allow_disable_backup is False
    assert cfg.cli_executable == ""


def test_load_config_without_file_uses_defaults(tmp_path: Path, monkeypatch) -> None:
    monkeypatch.delenv("HWPMATE_MCP_INPUT_ROOTS", raising=False)
    monkeypatch.delenv("HWPMATE_MCP_OUTPUT_ROOTS", raising=False)
    monkeypatch.delenv("HWPMATE_MCP_CLI", raising=False)
    cfg = load_config(tmp_path / "missing.toml")
    assert cfg.input_roots == [] and cfg.cli_executable == ""


def test_load_config_merges_toml_sections(tmp_path: Path, monkeypatch) -> None:
    monkeypatch.delenv("HWPMATE_MCP_INPUT_ROOTS", raising=False)
    monkeypatch.delenv("HWPMATE_MCP_OUTPUT_ROOTS", raising=False)
    monkeypatch.delenv("HWPMATE_MCP_CLI", raising=False)
    config = _write_toml(
        tmp_path / "config.toml",
        "[mcp]\nmax_files_per_job = 7\nallow_overwrite = true\n"
        "unknown_future_key = 1\n"
        '[paths]\ninput_roots = ["C:/qa-in"]\noutput_roots = ["C:/qa-out"]\n'
        'cli_executable = "C:/qa-cli.exe"\n',
    )
    cfg = load_config(config)
    assert cfg.max_files_per_job == 7
    assert cfg.allow_overwrite is True
    assert cfg.input_roots == ["C:/qa-in"]
    assert cfg.output_roots == ["C:/qa-out"]
    assert cfg.cli_executable == "C:/qa-cli.exe"


def test_env_overrides_toml_roots_and_cli(tmp_path: Path, monkeypatch) -> None:
    config = _write_toml(
        tmp_path / "config.toml",
        '[paths]\ninput_roots = ["C:/toml-in"]\ncli_executable = "C:/toml-cli.exe"\n',
    )
    monkeypatch.setenv("HWPMATE_MCP_INPUT_ROOTS", f"C:/env-a{os.pathsep}C:/env-b")
    monkeypatch.setenv("HWPMATE_MCP_OUTPUT_ROOTS", "C:/env-out")
    monkeypatch.setenv("HWPMATE_MCP_CLI", "C:/env-cli.exe")
    cfg = load_config(config)
    assert cfg.input_roots == ["C:/env-a", "C:/env-b"]
    assert cfg.output_roots == ["C:/env-out"]
    assert cfg.cli_executable == "C:/env-cli.exe"


def test_from_mapping_ignores_unknown_keys() -> None:
    cfg = McpConfig.from_mapping({"max_files_per_job": 9, "no_such_option": True})
    assert cfg.max_files_per_job == 9
    assert not hasattr(cfg, "no_such_option")


def test_effective_job_dir_prefers_explicit_setting(tmp_path: Path) -> None:
    cfg = McpConfig(job_dir=str(tmp_path / "jobs"))
    assert cfg.effective_job_dir() == tmp_path / "jobs"


def test_effective_job_dir_default_is_absolute() -> None:
    assert McpConfig().effective_job_dir().is_absolute()
    assert default_job_dir().is_absolute()
    assert default_config_path().is_absolute()
    assert default_config_path().name == "config.toml"
