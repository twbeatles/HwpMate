from __future__ import annotations

import time
from pathlib import Path
from typing import Iterable, List, Optional, Set

from PyQt6.QtCore import QThread, pyqtSignal

from ..constants import SCAN_BATCH_SIZE, SUPPORTED_EXTENSIONS
from ..logging_config import get_logger
from ..path_utils import canonicalize_path, iter_supported_files, make_path_key

logger = get_logger(__name__)

class FileScanWorker(QThread):
    """파일/폴더 입력을 비동기로 스캔하고 배치 단위로 전달"""

    batch_found = pyqtSignal(list)          # list[str]
    scan_progress = pyqtSignal(int, int)    # current_root, total_roots
    scan_finished = pyqtSignal(int, bool)   # total_found, canceled
    scan_error = pyqtSignal(str)

    def __init__(
        self,
        input_paths: List[str],
        include_sub: bool = True,
        allowed_exts: Optional[Iterable[str]] = None,
        batch_size: int = SCAN_BATCH_SIZE,
    ):
        super().__init__()
        self.input_paths = [str(p) for p in input_paths if str(p).strip()]
        self.include_sub = include_sub
        self.allowed_exts = {ext.lower() for ext in (allowed_exts or set(SUPPORTED_EXTENSIONS))}
        self.batch_size = max(1, int(batch_size))
        self._cancel_requested = False

    def cancel(self) -> None:
        self._cancel_requested = True

    def run(self) -> None:
        start_ts = time.perf_counter()
        batch: List[str] = []
        seen_keys: Set[str] = set()
        found_count = 0
        total_roots = len(self.input_paths)

        try:
            for idx, raw_path in enumerate(self.input_paths, start=1):
                if self._cancel_requested:
                    break

                root = Path(raw_path)
                for file_path in iter_supported_files(
                    root,
                    include_sub=self.include_sub,
                    allowed_exts=self.allowed_exts,
                    cancel_checker=lambda: self._cancel_requested,
                ):
                    if self._cancel_requested:
                        break

                    normalized = canonicalize_path(str(file_path))
                    key = make_path_key(normalized)
                    if key in seen_keys:
                        continue

                    seen_keys.add(key)
                    batch.append(normalized)
                    found_count += 1

                    if len(batch) >= self.batch_size:
                        self.batch_found.emit(batch)
                        batch = []

                self.scan_progress.emit(idx, total_roots)

            if batch:
                self.batch_found.emit(batch)

            elapsed = time.perf_counter() - start_ts
            logger.debug(
                f"FileScanWorker 완료: 입력={total_roots}, 발견={found_count}, "
                f"취소={self._cancel_requested}, 소요={elapsed:.3f}s"
            )
            self.scan_finished.emit(found_count, self._cancel_requested)
        except Exception as e:
            logger.exception("FileScanWorker 실행 중 오류")
            self.scan_error.emit(str(e))
