from __future__ import annotations

import os
import tempfile
from pathlib import Path

from PyQt6.QtCore import QLockFile

LOCK_FILE_NAME = "HwpMate.lock"


def default_lock_file_path() -> Path:
    local_app_data = os.environ.get("LOCALAPPDATA")
    if local_app_data:
        return Path(local_app_data) / "HwpMate" / LOCK_FILE_NAME

    home = Path.home()
    if str(home):
        return home / ".hwp_converter" / LOCK_FILE_NAME

    return Path(tempfile.gettempdir()) / "HwpMate" / LOCK_FILE_NAME


class SingleInstanceLock:
    """Small QLockFile wrapper for the GUI process lifetime."""

    def __init__(self, lock_file: Path | None = None) -> None:
        self.path = lock_file or default_lock_file_path()
        self.path.parent.mkdir(parents=True, exist_ok=True)
        self._lock = QLockFile(str(self.path))
        self._lock.setStaleLockTime(30_000)

    def try_lock(self) -> bool:
        return bool(self._lock.tryLock(0))

    def release(self) -> None:
        if self._lock.isLocked():
            self._lock.unlock()
