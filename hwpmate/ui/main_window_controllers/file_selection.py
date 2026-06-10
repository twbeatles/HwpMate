from __future__ import annotations

import logging
import time
from collections.abc import Iterable
from pathlib import Path
from typing import Any

from PyQt6.QtCore import QSignalBlocker
from PyQt6.QtWidgets import QFileDialog, QMessageBox, QTableWidgetItem

from ...constants import SCAN_BATCH_SIZE, SCAN_CANCEL_WAIT_MS, SUPPORTED_EXTENSIONS
from ...logging_config import get_logger
from ...path_utils import canonicalize_path
from ...workers.file_scan_worker import FileScanWorker
from .state import MainWindowState

logger = get_logger(__name__)


class FileSelectionController:
    """Folder/file selection and asynchronous scan lifecycle."""

    def __init__(self, window: Any, state: MainWindowState) -> None:
        self.window = window
        self.state = state

    def cancel_active_scan(self, wait_ms: int = SCAN_CANCEL_WAIT_MS) -> bool:
        worker = self.state.scan_worker
        if not worker:
            return True

        if worker.isRunning():
            worker.cancel()
            worker.wait(wait_ms)

        if worker.isRunning():
            return False

        try:
            worker.batch_found.disconnect(self.window._on_scan_batch_found)
            worker.scan_progress.disconnect(self.window._on_scan_progress)
            worker.scan_finished.disconnect(self.window._on_scan_finished)
            worker.scan_error.disconnect(self.window._on_scan_error)
            worker.finished.disconnect(self.window._on_scan_worker_finished)
        except (TypeError, RuntimeError):
            pass

        worker.deleteLater()
        self._clear_scan_state()
        return True

    def start_scan(
        self,
        input_paths: list[str],
        mode: str,
        include_sub: bool = True,
        allowed_exts: Iterable[str] | None = None,
    ) -> None:
        if self._input_locked():
            return

        cleaned_inputs = [str(p).strip() for p in input_paths if str(p).strip()]
        if not cleaned_inputs:
            return

        if not self.cancel_active_scan():
            logger.warning("이전 파일 스캔이 아직 종료되지 않아 새 스캔을 시작하지 않습니다.")
            return

        self.state.scan_mode = mode
        self.state.scan_new_file_count = 0
        self.state.scan_preview_count = 0
        self.state.scan_started_at = time.perf_counter()

        self.state.scan_worker = FileScanWorker(
            cleaned_inputs,
            include_sub=include_sub,
            allowed_exts=allowed_exts or SUPPORTED_EXTENSIONS,
            batch_size=SCAN_BATCH_SIZE,
        )
        self.state.scan_worker.batch_found.connect(self.window._on_scan_batch_found)
        self.state.scan_worker.scan_progress.connect(self.window._on_scan_progress)
        self.state.scan_worker.scan_finished.connect(self.window._on_scan_finished)
        self.state.scan_worker.scan_error.connect(self.window._on_scan_error)
        self.state.scan_worker.finished.connect(self.window._on_scan_worker_finished)
        self.state.scan_worker.start()

    def start_folder_preview_scan(self, folder_path: str) -> None:
        self.window.status_label.setText("📂 폴더 스캔 중...")
        self.start_scan(
            [folder_path],
            mode="folder_preview",
            include_sub=self.window.include_sub_check.isChecked(),
            allowed_exts=set(self.window.task_planner.preview_allowed_extensions(self.state.selected_format)),
        )

    def append_files_batch(self, files: list[str]) -> int:
        if not files:
            return 0

        unique_files = self.window.file_store.add_paths(files)
        if not unique_files:
            return 0

        render_start = time.perf_counter()
        start_row = self.window.file_table.rowCount()
        end_row = start_row + len(unique_files)

        self.window.file_table.setUpdatesEnabled(False)
        blocker = QSignalBlocker(self.window.file_table)
        try:
            self.window.file_table.setRowCount(end_row)
            for row_idx, file_path in enumerate(unique_files, start=start_row):
                file_obj = Path(file_path)
                self.window.file_table.setItem(row_idx, 0, QTableWidgetItem(file_obj.name))
                self.window.file_table.setItem(row_idx, 1, QTableWidgetItem(str(file_obj.parent)))
        finally:
            del blocker
            self.window.file_table.setUpdatesEnabled(True)

        self.update_file_count()

        if logger.isEnabledFor(logging.DEBUG):
            elapsed = time.perf_counter() - render_start
            logger.debug(f"파일 목록 렌더링: batch={len(unique_files)}, 소요={elapsed:.4f}s")
        return len(unique_files)

    def on_scan_batch_found(self, batch: list[str]) -> None:
        if self.window.sender() is not self.state.scan_worker:
            return

        if self.state.scan_mode == "add_files":
            added = self.append_files_batch(batch)
            self.state.scan_new_file_count += added
            return

        if self.state.scan_mode == "folder_preview":
            self.state.scan_preview_count += len(batch)

    def on_scan_progress(self, current: int, total: int) -> None:
        if self.window.sender() is not self.state.scan_worker:
            return

        if self.state.scan_mode == "add_files":
            self.window.status_label.setText(
                f"📥 파일 스캔 중... {current}/{total} 경로 처리 (신규 {self.state.scan_new_file_count}개)"
            )
            return

        if self.state.scan_mode == "folder_preview":
            self.window.status_label.setText(
                f"📂 폴더 스캔 중... {current}/{total} 경로 처리 ({self.state.scan_preview_count}개 발견)"
            )

    def on_scan_finished(self, total_found: int, canceled: bool) -> None:
        if self.window.sender() is not self.state.scan_worker:
            return

        elapsed = 0.0
        if self.state.scan_started_at is not None:
            elapsed = time.perf_counter() - self.state.scan_started_at

        if self.state.scan_mode == "add_files":
            if canceled:
                self.window.status_label.setText("파일 스캔이 취소되었습니다")
            elif self.state.scan_new_file_count == 0:
                self.window.status_label.setText("추가할 새 파일이 없습니다")
            else:
                self.window.status_label.setText(
                    f"{self.state.scan_new_file_count}개 파일 추가됨 (총 {len(self.window.file_list)}개)"
                )
            if logger.isEnabledFor(logging.DEBUG):
                logger.debug(
                    f"파일 추가 스캔 완료: 발견={total_found}, 신규={self.state.scan_new_file_count}, "
                    f"취소={canceled}, 소요={elapsed:.3f}s"
                )
            return

        if self.state.scan_mode == "folder_preview":
            if canceled:
                self.window.status_label.setText("폴더 스캔이 취소되었습니다")
            elif self.state.scan_preview_count == 0:
                self.window.status_label.setText("⚠️ 현재 포맷으로 변환 가능한 파일이 없습니다")
            else:
                self.window.status_label.setText(f"📁 {self.state.scan_preview_count}개 변환 가능 파일 발견")
            if logger.isEnabledFor(logging.DEBUG):
                logger.debug(
                    f"폴더 미리보기 스캔 완료: 발견={self.state.scan_preview_count}, "
                    f"취소={canceled}, 소요={elapsed:.3f}s"
                )

    def on_scan_error(self, error_msg: str) -> None:
        if self.window.sender() is not self.state.scan_worker:
            return
        logger.error(f"파일 스캔 오류: {error_msg}")
        self.window.status_label.setText("파일 스캔 중 오류가 발생했습니다")

    def on_scan_worker_finished(self) -> None:
        worker = self.state.scan_worker
        if self.window.sender() is not worker or worker is None:
            return
        worker.deleteLater()
        self._clear_scan_state()

    def select_folder(self) -> None:
        if self._input_locked():
            return
        initial = self.window.config.get("last_folder", "")
        folder = QFileDialog.getExistingDirectory(self.window, "폴더 선택", initial)
        if folder:
            self.window.folder_entry.setText(folder)
            self.window.config["last_folder"] = folder
            self.start_folder_preview_scan(folder)

    def select_output(self) -> None:
        if self._input_locked("변환 중에는 출력 폴더를 변경할 수 없습니다"):
            return
        initial = self.window.config.get("last_output", "")
        folder = QFileDialog.getExistingDirectory(self.window, "출력 폴더 선택", initial)
        if folder:
            self.window.output_entry.setText(folder)
            self.window.config["last_output"] = folder

    def browse_files(self) -> None:
        if self._input_locked():
            return
        files, _ = QFileDialog.getOpenFileNames(
            self.window,
            "파일 선택",
            "",
            "한글 파일 (*.hwp *.hwpx);;모든 파일 (*.*)",
        )
        if files:
            self.add_files(files)

    def add_files(self, files: list[str]) -> None:
        if self._input_locked():
            return
        if not files:
            return

        requested = [canonicalize_path(p) for p in files if str(p).strip()]
        if not requested:
            return

        scan_enqueue_start = time.perf_counter()
        self.window.status_label.setText(f"📥 {len(requested)}개 경로 스캔 시작...")
        self.start_scan(
            requested,
            mode="add_files",
            include_sub=True,
            allowed_exts=set(SUPPORTED_EXTENSIONS),
        )
        if logger.isEnabledFor(logging.DEBUG):
            elapsed = time.perf_counter() - scan_enqueue_start
            logger.debug(f"파일 스캔 요청 등록: 입력={len(requested)}, 소요={elapsed:.4f}s")

    def remove_selected(self) -> None:
        if self._input_locked("변환 중에는 파일 목록을 변경할 수 없습니다"):
            return
        selected = self.window.file_table.selectedItems()
        if not selected:
            return

        rows = set(item.row() for item in selected)
        self.window.file_store.remove_rows(rows)
        for row in sorted(rows, reverse=True):
            self.window.file_table.removeRow(row)

        self.window.status_label.setText(f"선택 파일 제거됨 (총 {len(self.window.file_list)}개)")
        self.update_file_count()

    def clear_all(self) -> None:
        if self._input_locked("변환 중에는 파일 목록을 변경할 수 없습니다"):
            return
        if not self.window.file_list:
            return

        reply = QMessageBox.question(
            self.window,
            "확인",
            f"{len(self.window.file_list)}개 파일을 모두 제거하시겠습니까?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        )

        if reply == QMessageBox.StandardButton.Yes:
            self.window.file_store.clear()
            self.window.file_table.setRowCount(0)
            self.window.status_label.setText("모든 파일 제거됨")
            self.update_file_count()

    def update_file_count(self) -> None:
        count = self.window.file_store.count
        self.window.file_count_label.setText(f"📄 파일: {count}개")

    def _clear_scan_state(self) -> None:
        self.state.scan_worker = None
        self.state.scan_mode = None
        self.state.scan_started_at = None
        self.state.scan_new_file_count = 0
        self.state.scan_preview_count = 0

    def _input_locked(self, message: str = "변환 중에는 입력을 변경할 수 없습니다") -> bool:
        worker = self.state.worker
        worker_running = bool(worker and getattr(worker, "isRunning", lambda: False)())
        if not (self.state.is_converting or worker_running):
            return False
        self.window.status_label.setText(message)
        if hasattr(self.window, "toast"):
            self.window.toast.show_message(message, "⚠️")
        return True
