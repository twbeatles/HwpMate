"""변환 전 원본 백업."""

from __future__ import annotations

import shutil
from datetime import datetime
from pathlib import Path

from ...constants import (
    BACKUP_MAX_FILES_PER_STEM,
    BACKUP_MAX_FILES_PER_STEM_MAX,
    BACKUP_MAX_FILES_PER_STEM_MIN,
)
from ...logging_config import get_logger

logger = get_logger(__name__)


def _clamp_backup_max(max_files: int | None) -> int:
    base = BACKUP_MAX_FILES_PER_STEM if max_files is None else int(max_files)
    return max(BACKUP_MAX_FILES_PER_STEM_MIN, min(BACKUP_MAX_FILES_PER_STEM_MAX, base))


def _prune_old_backups(
    backup_dir: Path,
    stem: str,
    suffix: str,
    *,
    keep_path: Path | None = None,
    max_files: int | None = None,
) -> None:
    """동일 stem(+suffix) 백업이 한도를 넘으면 오래된 것부터 삭제.

    방금 만든 keep_path 는 절대 삭제하지 않는다 (mtime 동점 오삭제 방지).
    """
    try:
        max_keep = _clamp_backup_max(max_files)
        prefix = f"{stem}_"
        keep_resolved = keep_path.resolve() if keep_path is not None else None
        candidates: list[Path] = []
        for entry in backup_dir.iterdir():
            if not entry.is_file():
                continue
            if entry.suffix.lower() != suffix.lower():
                continue
            if not entry.name.startswith(prefix):
                continue
            if keep_resolved is not None and entry.resolve() == keep_resolved:
                continue
            candidates.append(entry)

        # keep_path 1개를 포함한 총 상한
        slots_for_old = max_keep - (1 if keep_resolved is not None else 0)
        if slots_for_old < 0:
            slots_for_old = 0
        if len(candidates) <= slots_for_old:
            return
        candidates.sort(key=lambda p: (p.stat().st_mtime, p.name))
        for old in candidates[: len(candidates) - slots_for_old]:
            try:
                old.unlink(missing_ok=True)
                logger.debug(f"오래된 백업 정리: {old}")
            except OSError as e:
                logger.debug(f"백업 정리 실패(무시): {old} — {e}")
    except OSError as e:
        logger.debug(f"백업 정리 스캔 실패(무시): {e}")


def create_backup(file_path: Path, *, max_files: int | None = None) -> Path:
    """파일 백업 생성"""
    try:
        backup_dir = file_path.parent / "backup"
        backup_dir.mkdir(exist_ok=True)

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
        backup_name = f"{file_path.stem}_{timestamp}{file_path.suffix}"
        backup_path = backup_dir / backup_name
        counter = 1

        while backup_path.exists():
            backup_name = f"{file_path.stem}_{timestamp}_{counter}{file_path.suffix}"
            backup_path = backup_dir / backup_name
            counter += 1

        shutil.copy2(file_path, backup_path)
        logger.debug(f"백업 생성 완료: {backup_path}")
        _prune_old_backups(
            backup_dir,
            file_path.stem,
            file_path.suffix,
            keep_path=backup_path,
            max_files=max_files,
        )
        return backup_path
    except Exception as e:
        logger.error(f"백업 생성 중 오류: {e}")
        raise
