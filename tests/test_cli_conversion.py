from __future__ import annotations

import argparse
from pathlib import Path
from typing import Any
import pytest

from hwpmate.app import _parse_args, _run_cli_conversion


class MockCliConverter:
    def __init__(self) -> None:
        self.progid_used = "HWPControl.HwpCtrl.1"
        self.pdf_export_mode = "saveas_first"

    def initialize(self, *, manage_com_apartment: bool = True) -> bool:
        return True

    def convert_file(self, in_p: Any, out_p: Any, fmt: str = "PDF", *, cancel_check: Any = None) -> tuple[bool, str | None]:
        del cancel_check
        return True, None

    def cleanup(self) -> None:
        pass


class FakeLock:
    def __init__(self, available: bool = True) -> None:
        self.available = available
        self.released = False

    def try_lock(self) -> bool:
        return self.available

    def release(self) -> None:
        self.released = True


def test_parse_args_cli() -> None:
    args = _parse_args(["--input", "test.hwp", "--format", "DOCX", "--output", "out_dir", "--recursive", "--overwrite"])
    assert args.input == "test.hwp"
    assert args.format == "DOCX"
    assert args.output == "out_dir"
    assert args.recursive is True
    assert args.overwrite is True


def test_cli_conversion_file_not_found() -> None:
    args = _parse_args(["--input", "non_existent_file_path_xyz.hwp"])
    ret = _run_cli_conversion(args)
    assert ret == 1


def test_cli_conversion_invalid_format(tmp_path: Path) -> None:
    doc = tmp_path / "test.hwp"
    doc.write_bytes(b"dummy")
    args = _parse_args(["--input", str(doc), "--format", "INVALID_EXT"])
    ret = _run_cli_conversion(args)
    assert ret == 1


def test_cli_conversion_success(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    doc = tmp_path / "test.hwp"
    doc.write_bytes(b"dummy")

    import hwpmate.app as app_module
    monkeypatch.setattr(app_module, "PYWIN32_AVAILABLE", True)
    monkeypatch.setattr(app_module, "HWPConverter", MockCliConverter)
    monkeypatch.setattr(app_module, "SingleInstanceLock", lambda: FakeLock())

    args = _parse_args(["--input", str(doc), "--format", "PDF", "--output", str(tmp_path / "out")])
    ret = _run_cli_conversion(args)
    assert ret == 0


class ScriptedCliConverter(MockCliConverter):
    """호출마다 준비된 결과를 반환하고, 성공 시 산출물을 만든다."""

    instances: list["ScriptedCliConverter"] = []
    script: list[bool] = []

    def __init__(self) -> None:
        super().__init__()
        self.calls = 0
        self.auto_continue_compat_dialogs = True
        ScriptedCliConverter.instances.append(self)

    def convert_file(self, in_p: Any, out_p: Any, fmt: str = "PDF", *, cancel_check: Any = None) -> tuple[bool, str | None]:
        del in_p, fmt, cancel_check
        self.calls += 1
        ok = self.script.pop(0) if self.script else True
        if ok:
            Path(out_p).write_bytes(b"%PDF-1.4")
            return True, None
        return False, "temporary COM failure"


def _patch_cli(monkeypatch: pytest.MonkeyPatch, *, script: list[bool], lock: "FakeLock | None" = None) -> None:
    import hwpmate.app as app_module
    import hwpmate.workers.conversion_worker.task_runner as runner_module

    ScriptedCliConverter.instances = []
    ScriptedCliConverter.script = list(script)
    monkeypatch.setattr(app_module, "PYWIN32_AVAILABLE", True)
    monkeypatch.setattr(app_module, "HWPConverter", ScriptedCliConverter)
    monkeypatch.setattr(app_module, "SingleInstanceLock", lambda: lock or FakeLock())
    monkeypatch.setattr(runner_module, "RETRY_DELAY_SECONDS", 0)


def test_cli_applies_retry_count(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    doc = tmp_path / "a.hwp"
    doc.write_bytes(b"dummy")
    _patch_cli(monkeypatch, script=[False, False, True])

    ret = _run_cli_conversion(_parse_args(["--input", str(doc), "--retry", "2", "--no-backup"]))

    assert ret == 0
    assert ScriptedCliConverter.instances[0].calls == 3


def test_cli_creates_backup_by_default(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    doc = tmp_path / "a.hwp"
    doc.write_bytes(b"dummy")
    _patch_cli(monkeypatch, script=[True])

    ret = _run_cli_conversion(_parse_args(["--input", str(doc)]))

    assert ret == 0
    backups = list((tmp_path / "backup").glob("a_*.hwp"))
    assert len(backups) == 1


def test_cli_reports_skipped_same_format_and_writes_json_report(tmp_path: Path, monkeypatch: pytest.MonkeyPatch, capsys) -> None:
    import json

    (tmp_path / "x.hwp").write_bytes(b"dummy")
    (tmp_path / "y.hwpx").write_bytes(b"dummy")
    report = tmp_path / "out" / "result.json"
    _patch_cli(monkeypatch, script=[True])

    ret = _run_cli_conversion(
        _parse_args(["--input", str(tmp_path), "--format", "HWPX", "--no-backup", "--report", str(report)])
    )

    assert ret == 0
    assert "건너뜀 1건" in capsys.readouterr().out
    data = json.loads(report.read_text(encoding="utf-8"))
    assert data["summary"]["success_count"] == 1
    assert data["summary"]["skipped_count"] == 1


def test_cli_empty_folder_returns_error_without_exception(tmp_path: Path, monkeypatch: pytest.MonkeyPatch, capsys) -> None:
    _patch_cli(monkeypatch, script=[])

    ret = _run_cli_conversion(_parse_args(["--input", str(tmp_path)]))

    assert ret == 1
    assert "변환할 파일이 없습니다" in capsys.readouterr().err


def test_cli_rejects_non_hwp_input(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    doc = tmp_path / "a.docx"
    doc.write_bytes(b"dummy")
    _patch_cli(monkeypatch, script=[])

    assert _run_cli_conversion(_parse_args(["--input", str(doc)])) == 1
    assert ScriptedCliConverter.instances == []


def test_cli_refuses_when_single_instance_lock_is_held(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    doc = tmp_path / "a.hwp"
    doc.write_bytes(b"dummy")
    _patch_cli(monkeypatch, script=[True], lock=FakeLock(available=False))

    assert _run_cli_conversion(_parse_args(["--input", str(doc)])) == 1
    assert ScriptedCliConverter.instances == []


def test_cli_no_auto_continue_flag_disables_compat_dialog_response(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    doc = tmp_path / "a.hwp"
    doc.write_bytes(b"dummy")
    _patch_cli(monkeypatch, script=[True])

    _run_cli_conversion(_parse_args(["--input", str(doc), "--no-backup", "--no-auto-continue"]))

    assert ScriptedCliConverter.instances[0].auto_continue_compat_dialogs is False
