from __future__ import annotations

from pathlib import Path

import pytest

from hwpmate.mcp.config import McpConfig
from hwpmate.mcp.errors import McpError
from hwpmate.mcp.runner import build_cli_argv, is_busy_failure, resolve_cli_executable, run_cli


def test_resolve_cli_requires_explicit_file(tmp_path: Path) -> None:
    with pytest.raises(McpError) as excinfo:
        resolve_cli_executable(McpConfig())
    assert excinfo.value.code == "CLI_NOT_TRUSTED"
    with pytest.raises(McpError) as excinfo:
        resolve_cli_executable(McpConfig(cli_executable=str(tmp_path / "nope.exe")))
    assert excinfo.value.code == "CLI_NOT_TRUSTED"


def test_resolve_py_uses_allowed_interpreter(tmp_path: Path) -> None:
    script = tmp_path / "cli.py"
    script.write_text("x", encoding="utf-8")
    base, display = resolve_cli_executable(McpConfig(cli_executable=str(script)))
    assert base[0].endswith("python.exe") or base[0] == "python" or "python" in base[0].lower()
    assert base[1] == str(script.resolve())
    assert display == str(script.resolve())


def test_argv_is_fixed_array() -> None:
    argv = build_cli_argv(
        ["exe"],
        input_path="a.hwp",
        format_type="PDF",
        output_dir="out",
        report_path="r.json",
    )
    assert argv == [
        "exe", "--input", "a.hwp", "--format", "PDF", "--output", "out",
        "--report", "r.json", "--retry", "1", "--pdf-export-mode", "saveas_first",
    ]
    assert "--overwrite" not in argv
    assert "--no-backup" not in argv


def test_run_cli_captures_logs_to_files(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    fake = Path(__file__).parent / "fake_cli.py"
    monkeypatch.setenv("HWPMATE_FAKECLI_MODE", "ok")
    cfg = McpConfig(cli_executable=str(fake))
    base, _ = resolve_cli_executable(cfg)
    argv = build_cli_argv(
        base, input_path="a.hwp", format_type="PDF",
        output_dir=str(tmp_path / "out"), report_path=str(tmp_path / "r.json"),
    )
    run = run_cli(cfg, argv=argv, job_private_dir=tmp_path / "job", report_path=tmp_path / "r.json", timeout_seconds=60)
    assert run.returncode == 0
    assert not is_busy_failure(run)
    assert Path(run.stdout_path).is_file()
    assert Path(run.stderr_path).is_file()


def test_run_cli_timeout(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    fake = Path(__file__).parent / "fake_cli.py"
    monkeypatch.setenv("HWPMATE_FAKECLI_MODE", "sleep")
    monkeypatch.setenv("HWPMATE_FAKECLI_SLEEP", "30")
    cfg = McpConfig(cli_executable=str(fake))
    base, _ = resolve_cli_executable(cfg)
    argv = build_cli_argv(
        base, input_path="a.hwp", format_type="PDF",
        output_dir=str(tmp_path / "out"), report_path=str(tmp_path / "r.json"),
    )
    run = run_cli(cfg, argv=argv, job_private_dir=tmp_path / "job", report_path=tmp_path / "r.json", timeout_seconds=2)
    assert run.timed_out is True
    assert run.returncode == 124
