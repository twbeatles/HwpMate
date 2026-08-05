"""변환 전 원본 백업."""

from __future__ import annotations

import shutil
from datetime import datetime
from pathlib import Path

from ...logging_config import get_logger

logger = get_logger(__name__)


def create_backup(file_path: Path) -> Path:
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
        return backup_path
    except Exception as e:
        logger.error(f"백업 생성 중 오류: {e}")
        raise
