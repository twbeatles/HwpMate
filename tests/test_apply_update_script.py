from __future__ import annotations

import hashlib
import json
from pathlib import Path
import pytest

import scripts.apply_update as apply_update_module
from scripts.apply_update import main


def test_apply_update_script_success(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    target = tmp_path / "app.exe"
    staged = tmp_path / "staged.exe"
    backup = tmp_path / "app.exe.v9.0.bak"
    result_file = tmp_path / "result.json"

    target.write_bytes(b"v9.0")
    new_bytes = b"v9.1"
    staged.write_bytes(new_bytes)

    monkeypatch.setattr(apply_update_module, "_wait_for_parent", lambda *args: None)
    monkeypatch.setattr(apply_update_module, "apply_staged_update", lambda **kwargs: None)

    ret = main([
        "--target", str(target),
        "--staged", str(staged),
        "--backup", str(backup),
        "--parent-pid", "1",
        "--expected-sha256", hashlib.sha256(new_bytes).hexdigest(),
        "--expected-size", str(len(new_bytes)),
        "--result-file", str(result_file),
    ])

    assert ret == 0
    assert result_file.is_file()
    data = json.loads(result_file.read_text(encoding="utf-8"))
    assert data["status"] == "applied"


def test_apply_update_script_records_rolled_back_status_from_exception(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    from hwpmate.services.update_installer import UpdateApplyError

    result_file = tmp_path / "result.json"

    def fail_with_rollback(**kwargs):
        raise UpdateApplyError("smoke failed", status="rolled_back")

    monkeypatch.setattr(apply_update_module, "_wait_for_parent", lambda *args: None)
    monkeypatch.setattr(apply_update_module, "wait_for_file_writable", lambda *args, **kwargs: True)
    monkeypatch.setattr(apply_update_module, "apply_staged_update", fail_with_rollback)

    ret = main([
        "--target", str(tmp_path / "app.exe"),
        "--staged", str(tmp_path / "staged.exe"),
        "--backup", str(tmp_path / "app.exe.v9.0.bak"),
        "--parent-pid", "1",
        "--expected-sha256", "0" * 64,
        "--expected-size", "1",
        "--result-file", str(result_file),
    ])

    assert ret == 1
    assert json.loads(result_file.read_text(encoding="utf-8"))["status"] == "rolled_back"


def test_apply_update_script_does_not_misreport_rollback_failure(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    from hwpmate.services.update_installer import UpdateApplyError

    result_file = tmp_path / "result.json"

    def fail_rollback(**kwargs):
        raise UpdateApplyError("업데이트 적용과 롤백 복구가 모두 실패했습니다.", status="failed")

    monkeypatch.setattr(apply_update_module, "_wait_for_parent", lambda *args: None)
    monkeypatch.setattr(apply_update_module, "wait_for_file_writable", lambda *args, **kwargs: True)
    monkeypatch.setattr(apply_update_module, "apply_staged_update", fail_rollback)

    main([
        "--target", str(tmp_path / "app.exe"),
        "--staged", str(tmp_path / "staged.exe"),
        "--backup", str(tmp_path / "app.exe.v9.0.bak"),
        "--parent-pid", "1",
        "--expected-sha256", "0" * 64,
        "--expected-size", "1",
        "--result-file", str(result_file),
    ])

    assert json.loads(result_file.read_text(encoding="utf-8"))["status"] == "failed"


def test_app_apply_update_relaunches_after_rollback(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    import hwpmate.app as app_module
    from hwpmate.services.update_installer import UpdateApplyError

    relaunched: list[Path] = []

    def fail_with_rollback(**kwargs):
        raise UpdateApplyError("smoke failed", status="rolled_back")

    monkeypatch.setattr(app_module, "_wait_for_parent", lambda *args: None)
    monkeypatch.setattr(app_module, "wait_for_file_writable", lambda *args, **kwargs: True)
    monkeypatch.setattr(app_module, "apply_staged_update", fail_with_rollback)
    monkeypatch.setattr(app_module, "_relaunch", lambda target: relaunched.append(target))
    result_file = tmp_path / "result.json"
    args = app_module._parse_args([
        "--apply-update",
        "--update-target", str(tmp_path / "app.exe"),
        "--update-staged", str(tmp_path / "staged.exe"),
        "--update-backup", str(tmp_path / "app.exe.v9.0.bak"),
        "--update-parent-pid", "1",
        "--update-result-file", str(result_file),
    ])

    assert app_module._run_apply_update(args) == 1
    assert relaunched == [tmp_path / "app.exe"]
    assert json.loads(result_file.read_text(encoding="utf-8"))["status"] == "rolled_back"
