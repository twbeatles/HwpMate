"""P0-B — 백업 생성·보존·정리 격리 회귀 테스트.

서로 접두사가 겹치는 stem(report / report_1 / report_final)의 백업이
한 stem 정리에 침범하지 않는지 파일 목록·바이트 해시로 증명한다.
백업 실패 정책(변환 계속 + 명시 기록)이 거짓 성공으로 보이지 않는지도 검증한다.
"""

from __future__ import annotations

import hashlib
from pathlib import Path

from hwpmate.constants import BACKUP_DIR_NAME, BACKUP_MAX_FILES_PER_STEM
from hwpmate.models import ConversionTask
from hwpmate.services.task_planner import TaskPlanner
from hwpmate.workers.conversion_worker.backup import create_backup
from hwpmate.workers.conversion_worker.task_runner import TaskRunOptions, execute_task


def _sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


class OkConverter:
    """성공 + 산출물 기록을 남기는 충실한 fake 엔진 (ConverterEngine 충족)."""

    progid_used: str | None = "Fake.Hwp"
    pdf_export_mode = "saveas_first"
    last_created_files: list[Path] = []

    def initialize(self, *, manage_com_apartment: bool = True) -> bool:
        del manage_com_apartment
        return True

    def cleanup(self) -> None:
        return None

    def has_owned_processes(self) -> bool:
        return False

    def kill_owned_processes(self) -> bool:
        return False

    def convert_file(self, input_path, output_path, format_type="PDF", *, cancel_check=None):
        del input_path, format_type, cancel_check
        Path(output_path).write_bytes(b"%PDF-1.4 ok\n")
        return True, None


def _snapshot_hashes(directory: Path) -> dict[str, str]:
    return {
        entry.name: _sha256(entry)
        for entry in sorted(directory.iterdir())
        if entry.is_file()
    }


def test_prune_never_touches_other_stems_with_overlapping_prefix(tmp_path: Path) -> None:
    """report 정리 시 report_1·report_final 백업은 개수·내용·해시 모두 불변."""
    report = tmp_path / "report.hwp"
    report.write_text("v", encoding="utf-8")

    # 겹치는 접두사의 이웃 stem 백업을 먼저 만든다.
    neighbors = {
        "report_1": tmp_path / "report_1.hwp",
        "report_final": tmp_path / "report_final.hwpx",
    }
    for stem, source in neighbors.items():
        source.write_text(f"origin-{stem}", encoding="utf-8")
        create_backup(source)
        create_backup(source)
    backup_dir = tmp_path / BACKUP_DIR_NAME
    neighbor_hashes = {
        stem: _snapshot_hashes(backup_dir) for stem in ("report_1", "report_final")
    }
    neighbor_counts = {
        stem: sum(1 for name in hashes if name.startswith(f"{stem}_"))
        for stem, hashes in neighbor_hashes.items()
    }

    # report stem 백업을 한도 초과로 쌓아 정리를 유발한다.
    keep = 3
    for _ in range(BACKUP_MAX_FILES_PER_STEM + keep + 2):
        create_backup(report, max_files=keep)

    after = _snapshot_hashes(backup_dir)
    report_backups = [name for name in after if name.startswith("report_20")]
    assert 0 < len(report_backups) <= keep
    for stem, before in neighbor_hashes.items():
        stem_files = {name: digest for name, digest in after.items() if name.startswith(f"{stem}_")}
        expected = {name: digest for name, digest in before.items() if name.startswith(f"{stem}_")}
        assert stem_files == expected, stem
        assert len(stem_files) == neighbor_counts[stem], stem


def test_prune_orders_by_backup_name_not_source_mtime(tmp_path: Path) -> None:
    """원본 mtime 이 오래됐다고 최근 백업을 먼저 삭제하면 안 된다."""
    old_source = tmp_path / "doc.hwp"
    old_source.write_bytes(b"old")
    first = create_backup(old_source, max_files=10)

    new_source = tmp_path / "doc.hwp"
    new_source.write_bytes(b"new-content")
    second = create_backup(new_source, max_files=10)
    assert first != second

    # 한도를 1개로 조여 가장 오래된 백업 1개만 남긴다.
    create_backup(new_source, max_files=1)
    backup_dir = tmp_path / BACKUP_DIR_NAME
    remaining = [entry for entry in backup_dir.iterdir() if entry.is_file()]
    assert len(remaining) == 1
    assert remaining[0].read_bytes() == b"new-content"


def test_backup_failure_is_recorded_and_warned_not_silent(tmp_path: Path) -> None:
    """백업 실패는 변환을 계속하되 task·경고에 명시된다 (거짓 성공 금지)."""

    def _boom(path: Path) -> Path:
        raise PermissionError("denied by policy")

    src = tmp_path / "a.hwp"
    src.write_bytes(b"dummy-hwp")
    task = ConversionTask(input_file=src, output_file=tmp_path / "a.pdf")

    warnings = execute_task(
        task,
        OkConverter(),
        TaskRunOptions(format_type="PDF", overwrite=False, backup_enabled=True, retry_count=0),
        planner=TaskPlanner(),
        used_output_path_keys=set(),
        cancel_check=lambda: False,
        create_backup_fn=_boom,
    )
    assert task.status == "성공"
    assert task.output_file.is_file()
    assert task.backup_error == "denied by policy"
    assert task.backup_file is None
    assert any("백업 실패" in warning for warning in warnings)


def test_backup_dir_swapped_by_outsider_is_reported(tmp_path: Path) -> None:
    """백업 폴더 자리가 외부 파일로 교체돼도 원본·변환은 보호되고 기록된다."""
    src = tmp_path / "a.hwp"
    src.write_bytes(b"dummy-hwp")
    (tmp_path / BACKUP_DIR_NAME).write_bytes(b"not-a-directory")

    backup_path = None
    error: str | None = None
    try:
        backup_path = create_backup(src)
    except Exception as exc:  # noqa: BLE001
        error = str(exc)
    assert backup_path is None
    assert error, "백업 폴더 교체는 명확한 오류여야 한다"
    assert src.read_bytes() == b"dummy-hwp"

    task = ConversionTask(input_file=src, output_file=tmp_path / "a.pdf")

    warnings = execute_task(
        task,
        OkConverter(),
        TaskRunOptions(format_type="PDF", overwrite=False, backup_enabled=True, retry_count=0),
        planner=TaskPlanner(),
        used_output_path_keys=set(),
        cancel_check=lambda: False,
    )
    assert task.status == "성공"
    assert task.backup_error, "백업 실패는 task 에 기록되어야 한다"
    assert any("백업 실패" in warning for warning in warnings)
