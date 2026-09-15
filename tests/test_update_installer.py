from __future__ import annotations

import hashlib
from datetime import datetime, timedelta, timezone
from pathlib import Path

import pytest

from hwpmate.services.update_installer import (
    UpdateApplyError,
    apply_staged_update,
    cleanup_update_backups,
    consume_update_result,
    prepare_staged_update,
    write_update_result,
)
from hwpmate.services.update_manifest import ReleaseManifest


@pytest.fixture
def sample_manifest() -> ReleaseManifest:
    content = b"fake executable binary content for testing"
    return ReleaseManifest(
        version="9.1.0",
        artifact_url="https://example.com/app.exe",
        artifact_sha256=hashlib.sha256(content).hexdigest(),
        artifact_size=len(content),
        expires_at=datetime.now(timezone.utc) + timedelta(days=1),
        signature="fake-signature",
    )


def test_prepare_staged_update_success(sample_manifest: ReleaseManifest, tmp_path: Path) -> None:
    chunks = [b"fake executable ", b"binary content ", b"for testing"]
    staged = prepare_staged_update(
        sample_manifest,
        chunks=chunks,
        staging_root=tmp_path,
        approve=lambda m, p: True,
    )
    assert staged is not None
    assert staged.is_file()
    assert staged.read_bytes() == b"fake executable binary content for testing"


def test_prepare_staged_update_hash_mismatch(sample_manifest: ReleaseManifest, tmp_path: Path) -> None:
    chunks = [b"corrupted binary payload"]
    with pytest.raises(ValueError, match="일치하지 않습니다"):
        prepare_staged_update(
            sample_manifest,
            chunks=chunks,
            staging_root=tmp_path,
        )
    # 임시 파일이 정리되었는지 확인
    assert len(list(tmp_path.glob("update-*.exe"))) == 0


def test_prepare_staged_update_user_declined(sample_manifest: ReleaseManifest, tmp_path: Path) -> None:
    chunks = [b"fake executable binary content for testing"]
    staged = prepare_staged_update(
        sample_manifest,
        chunks=chunks,
        staging_root=tmp_path,
        approve=lambda m, p: False,
    )
    assert staged is None
    assert len(list(tmp_path.glob("update-*.exe"))) == 0


def test_apply_staged_update_success_and_smoke_pass(tmp_path: Path) -> None:
    target = tmp_path / "app.exe"
    staged = tmp_path / "staged.exe"
    backup = tmp_path / "app.exe.v9.0.bak"

    old_bytes = b"version 9.0"
    new_bytes = b"version 9.1"
    target.write_bytes(old_bytes)
    staged.write_bytes(new_bytes)

    apply_staged_update(
        target=target,
        staged=staged,
        backup=backup,
        expected_sha256=hashlib.sha256(new_bytes).hexdigest(),
        expected_size=len(new_bytes),
        smoke_runner=lambda p: True,
    )

    assert target.read_bytes() == new_bytes
    assert backup.read_bytes() == old_bytes


def test_apply_staged_update_rollback_on_smoke_failure(tmp_path: Path) -> None:
    target = tmp_path / "app.exe"
    staged = tmp_path / "staged.exe"
    backup = tmp_path / "app.exe.v9.0.bak"

    old_bytes = b"version 9.0"
    bad_bytes = b"broken binary"
    target.write_bytes(old_bytes)
    staged.write_bytes(bad_bytes)

    with pytest.raises(UpdateApplyError, match="롤백됨"):
        apply_staged_update(
            target=target,
            staged=staged,
            backup=backup,
            expected_sha256=hashlib.sha256(bad_bytes).hexdigest(),
            expected_size=len(bad_bytes),
            smoke_runner=lambda p: False,  # 스모크 실패 시뮬레이션
        )

    # 타깃 파일이 이전 버전(9.0)으로 복구되었는지 확인
    assert target.read_bytes() == old_bytes


def test_write_and_consume_update_result(tmp_path: Path) -> None:
    res_path = tmp_path / "last-update-result.json"
    write_update_result(res_path, {"status": "applied", "version": "9.1.0"})
    assert res_path.is_file()

    result = consume_update_result(res_path)
    assert result is not None
    assert result["status"] == "applied"
    assert result["version"] == "9.1.0"

    # 소비 후 삭제되었는지 확인
    assert not res_path.exists()
    assert consume_update_result(res_path) is None


def test_cleanup_update_backups(tmp_path: Path) -> None:
    target = tmp_path / "app.exe"
    target.write_text("main")

    b1 = tmp_path / "app.exe.v1.bak"
    b2 = tmp_path / "app.exe.v2.bak"
    b3 = tmp_path / "app.exe.v3.bak"
    b4 = tmp_path / "app.exe.v4.bak"

    b1.write_text("1")
    b2.write_text("2")
    b3.write_text("3")
    b4.write_text("4")

    cleanup_update_backups(target, keep_count=2)
    remaining = sorted(p.name for p in tmp_path.glob("app.exe.v*.bak"))
    assert len(remaining) == 2


def test_apply_staged_update_replace_failure_keeps_old_version_and_removes_backup(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    import hwpmate.services.update_installer as installer

    target = tmp_path / "app.exe"
    staged = tmp_path / "staged.exe"
    backup = tmp_path / "app.exe.v9.0.bak"
    target.write_bytes(b"old")
    staged.write_bytes(b"new")

    def locked_replace(src, dst):
        raise PermissionError("target locked")

    monkeypatch.setattr(installer.os, "replace", locked_replace)

    with pytest.raises(UpdateApplyError) as exc_info:
        apply_staged_update(
            target=target,
            staged=staged,
            backup=backup,
            smoke_runner=lambda p: True,
            replace_attempts=2,
            replace_delay=0,
        )

    assert exc_info.value.status == "failed"
    assert target.read_bytes() == b"old"
    assert not backup.exists()


def test_apply_staged_update_rollback_status_is_explicit(tmp_path: Path) -> None:
    target = tmp_path / "app.exe"
    staged = tmp_path / "staged.exe"
    backup = tmp_path / "app.exe.v9.0.bak"
    target.write_bytes(b"old")
    staged.write_bytes(b"bad")

    with pytest.raises(UpdateApplyError) as exc_info:
        apply_staged_update(target=target, staged=staged, backup=backup, smoke_runner=lambda p: False)

    assert exc_info.value.status == "rolled_back"
    assert target.read_bytes() == b"old"


def test_apply_staged_update_rollback_failure_is_reported_as_failed(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    import hwpmate.services.update_installer as installer

    target = tmp_path / "app.exe"
    staged = tmp_path / "staged.exe"
    backup = tmp_path / "app.exe.v9.0.bak"
    target.write_bytes(b"old")
    staged.write_bytes(b"bad")
    real_replace = installer.os.replace
    calls: list[tuple[str, str]] = []

    def replace_then_lock(src, dst):
        calls.append((str(src), str(dst)))
        if len(calls) == 1:
            return real_replace(src, dst)
        raise PermissionError("locked during rollback")

    monkeypatch.setattr(installer.os, "replace", replace_then_lock)

    with pytest.raises(UpdateApplyError) as exc_info:
        apply_staged_update(
            target=target,
            staged=staged,
            backup=backup,
            smoke_runner=lambda p: False,
            replace_attempts=2,
            replace_delay=0,
        )

    assert exc_info.value.status == "failed"
    assert "수동 복구" in str(exc_info.value)
    assert backup.read_bytes() == b"old"  # 유일한 정상본은 남긴다


def test_unique_update_backup_path_skips_existing_backups(tmp_path: Path) -> None:
    from hwpmate.services.update_installer import unique_update_backup_path

    target = tmp_path / "app.exe"
    target.write_bytes(b"x")
    (tmp_path / "app.exe.v9.1.0.bak").write_bytes(b"old")
    (tmp_path / "app.exe.v9.1.0.1.bak").write_bytes(b"old")

    assert unique_update_backup_path(target, "9.1.0") == (tmp_path / "app.exe.v9.1.0.2.bak").resolve()


def test_cleanup_update_staging_removes_helpers_and_stale_downloads(tmp_path: Path) -> None:
    from hwpmate.services.update_installer import cleanup_update_staging

    helper = tmp_path / "update-helper-abc.exe"
    stale = tmp_path / "update-9.2.0-def.exe"
    keep = tmp_path / "update-9.3.0-ghi.exe"
    result = tmp_path / "last-update-result.json"
    for path in (helper, stale, keep, result):
        path.write_bytes(b"x")

    removed = cleanup_update_staging(tmp_path, keep=[keep])

    assert removed == 2
    assert not helper.exists() and not stale.exists()
    assert keep.exists() and result.exists()


def test_wait_for_process_exit_really_waits_for_live_process() -> None:
    import subprocess
    import sys

    from hwpmate.services.update_installer import wait_for_process_exit

    child = subprocess.Popen([sys.executable, "-c", "import time; time.sleep(30)"])
    try:
        # os.kill(pid, 0) 기반 구현은 콘솔 없는 헬퍼에서 즉시 반환해 버렸다.
        assert wait_for_process_exit(child.pid, 0.3) is False
    finally:
        child.kill()
        child.wait(timeout=10)
    assert wait_for_process_exit(child.pid, 5.0) is True


def test_wait_for_process_exit_accepts_missing_pid() -> None:
    from hwpmate.services.update_installer import wait_for_process_exit

    assert wait_for_process_exit(0) is True
