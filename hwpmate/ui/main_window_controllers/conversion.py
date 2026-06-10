from __future__ import annotations

import logging
import time
from pathlib import Path
from typing import Any

from PyQt6.QtCore import QTimer
from PyQt6.QtWidgets import QApplication, QMessageBox

from ...constants import WORKER_WAIT_TIMEOUT
from ...logging_config import get_logger
from ...models import ConversionSummary, PlannedConversion
from ...path_utils import check_write_permission, is_valid_path_name
from .state import MainWindowState

logger = get_logger(__name__)


class ConversionController:
    """Task planning, worker orchestration, and conversion result handling."""

    def __init__(self, window: Any, state: MainWindowState) -> None:
        self.window = window
        self.state = state

    def collect_tasks(self) -> PlannedConversion:
        return self.window.task_planner.build_tasks(
            is_folder_mode=self.window.folder_radio.isChecked(),
            format_type=self.state.selected_format,
            folder_path=self.window.folder_entry.text(),
            include_sub=self.window.include_sub_check.isChecked(),
            same_location=self.window.same_location_check.isChecked(),
            output_path=self.window.output_entry.text(),
            file_paths=self.window.file_store.paths,
            backup_enabled=self.window.backup_check.isChecked(),
            retry_count=self.window.retry_spin.value(),
        )

    def adjust_output_paths(self, plan: PlannedConversion, *, overwrite: bool) -> int:
        return self.window.task_planner.resolve_output_conflicts(plan.tasks, overwrite=overwrite)

    def validate_output_settings(self) -> None:
        if self.window.same_location_check.isChecked():
            return

        output_path = self.window.output_entry.text().strip()
        if not output_path:
            raise ValueError("출력 폴더를 선택하세요.")
        if not is_valid_path_name(output_path):
            raise ValueError(f"출력 경로에 사용할 수 없는 문자가 있습니다:\n{output_path}")

        output_folder = Path(output_path)
        if not output_folder.exists():
            raise ValueError(f"출력 폴더가 존재하지 않습니다:\n{output_folder}")
        if not check_write_permission(output_folder):
            raise ValueError(f"출력 폴더에 쓰기 권한이 없습니다:\n{output_folder}")

    def start_conversion(self) -> None:
        try:
            if self.state.scan_worker and self.state.scan_worker.isRunning():
                if self.state.scan_mode == "add_files":
                    QMessageBox.warning(self.window, "경고", "파일 스캔이 진행 중입니다. 스캔 완료 후 다시 시도하세요.")
                    return
                if not self.window._cancel_active_scan():
                    QMessageBox.warning(self.window, "경고", "폴더 스캔이 아직 종료되지 않았습니다. 잠시 후 다시 시도하세요.")
                    return

            self.validate_output_settings()

            task_collect_start = time.perf_counter()
            plan = self.collect_tasks()
            if logger.isEnabledFor(logging.DEBUG):
                logger.debug(
                    f"작업 목록 생성 완료: 실행={plan.runnable_count}개, 건너뜀={plan.skipped_count}개, "
                    f"소요={time.perf_counter() - task_collect_start:.3f}s"
                )

            overwrite = self.window.overwrite_check.isChecked()
            plan.conflict_renamed_count = self.adjust_output_paths(plan, overwrite=overwrite)
            if plan.conflict_renamed_count:
                if overwrite:
                    plan.warnings.append(
                        f"실행 배치 내부 출력 경로 충돌 {plan.conflict_renamed_count}개는 자동으로 새 이름으로 저장됩니다."
                    )
                else:
                    plan.warnings.append(
                        f"출력 경로 충돌 {plan.conflict_renamed_count}개는 자동으로 새 이름으로 저장됩니다."
                    )

            if not plan.tasks and plan.skipped_count:
                self.state.plan = plan
                self.window._save_settings()
                self.show_skipped_only_result(plan)
                return

            if not plan.tasks:
                message = "실행할 변환 대상이 없습니다."
                if plan.skipped_count:
                    message += f"\n동일 형식 {plan.skipped_count}개는 자동으로 건너뜁니다."
                raise ValueError(message)

            preflight = self.window._create_preflight_dialog(plan)
            if preflight.exec() != self.window.dialog_accepted_code():
                self.window.status_label.setText("변환 시작이 취소되었습니다")
                return

            self.state.plan = plan
            self.state.tasks = plan.tasks
            self.window._save_settings()

            self.window._set_converting_state(True)
            self.window.progress_bar.setMaximum(plan.runnable_count)
            self.window.progress_bar.setValue(0)
            self.state.conversion_start_time = time.time()
            worker = self.window._create_conversion_worker(plan)
            self.state.worker = worker
            worker.progress_updated.connect(self.window._on_progress_updated)
            worker.status_updated.connect(self.window._on_status_updated)
            worker.task_completed.connect(self.window._on_task_completed)
            worker.error_occurred.connect(self.window._on_error_occurred)
            worker.finished.connect(self.window._on_worker_finished)
            worker.start()
            self.window.hwp_status_label.setText("🟡 한글 연결 중...")

            start_message = f"{plan.runnable_count}개 파일 변환 시작"
            if plan.skipped_count:
                start_message += f" (건너뜀 {plan.skipped_count}개)"
            self.window.toast.show_message(start_message, "🚀")
        except ValueError as e:
            QMessageBox.warning(self.window, "경고", str(e))
        except Exception as e:
            logger.exception("변환 시작 오류")
            QMessageBox.critical(self.window, "오류", f"오류 발생: {e}")

    def show_skipped_only_result(self, plan: PlannedConversion) -> None:
        summary = ConversionSummary(
            format_type=plan.format_type,
            tasks=list(plan.skipped_tasks),
            warnings=list(plan.warnings),
            elapsed_seconds=0.0,
        )
        self.state.last_summary = summary
        self.window.status_label.setText("동일 형식 파일만 있어 변환 없이 건너뜀 처리했습니다")
        self.window.toast.show_message(f"건너뜀 {summary.skipped_count}개", "⏭️")
        dialog = self.window._create_result_dialog(summary)
        dialog.exec()
        self.state.plan = None

    def request_worker_stop(self, waiting_text: str) -> bool:
        worker = self.state.worker
        if worker is None:
            return True

        self.window.status_label.setText(waiting_text)
        worker.cancel()
        if worker.wait(WORKER_WAIT_TIMEOUT):
            return True

        if worker.can_force_terminate():
            self.state.force_kill_pending = True
            self.window.cancel_btn.setText("🛑 강제 종료")
            self.window.status_label.setText("취소 요청됨 (응답 대기)")
        else:
            self.state.force_kill_pending = False
            self.window.cancel_btn.setText("⏹️ 취소")
            self.window.status_label.setText("안전하게 강제 종료할 대상 프로세스를 확인하지 못했습니다. 종료를 기다리는 중입니다.")
        return False

    def perform_force_terminate(self) -> bool:
        worker = self.state.worker
        if worker is None:
            return False

        self.window.status_label.setText("강제 종료 중...")
        QApplication.processEvents()
        killed = worker.force_terminate()
        if not killed:
            self.state.force_kill_pending = False
            self.window.cancel_btn.setText("⏹️ 취소")
            QMessageBox.warning(
                self.window,
                "강제 종료 불가",
                "안전하게 종료할 대상 프로세스를 확인하지 못해 강제 종료를 수행하지 않았습니다.",
            )
            self.window.status_label.setText("안전한 강제 종료 대상이 없어 종료를 기다리는 중입니다.")
            return False

        worker.wait(1000)
        self.state.force_kill_pending = False
        self.window.cancel_btn.setText("⏹️ 취소")
        return True

    def cancel_conversion(self) -> None:
        if not self.state.worker:
            return

        if self.state.force_kill_pending:
            reply = QMessageBox.question(
                self.window,
                "강제 종료 경고",
                "앱이 소유한 한글 프로세스만 강제 종료합니다.\n열려 있는 문서가 저장되지 않을 수 있습니다.\n\n계속할까요?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            )
            if reply == QMessageBox.StandardButton.Yes:
                if self.perform_force_terminate():
                    self.window.status_label.setText("강제 종료 요청 완료")
            return

        reply = QMessageBox.question(
            self.window,
            "확인",
            "변환을 취소하시겠습니까?\n응답이 없으면 앱이 소유한 한글 프로세스만 강제 종료할 수 있습니다.",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        )
        if reply != QMessageBox.StandardButton.Yes:
            return

        if self.request_worker_stop("취소 요청 중..."):
            self.window.status_label.setText("취소됨")

    def on_progress_updated(self, current: int, total: int, filename: str) -> None:
        self.window.progress_bar.setValue(current)

        if current > 0 and self.state.conversion_start_time:
            elapsed = time.time() - self.state.conversion_start_time
            avg_time = elapsed / current
            remaining = avg_time * (total - current)
            remaining_str = f" (남은 시간: {int(remaining)}초)" if remaining > 0 else ""
        else:
            remaining_str = ""

        self.window.progress_label.setText(f"{current} / {total}{remaining_str}")
        self.window.status_label.setText(f"변환 중: {filename}")

    def on_status_updated(self, text: str) -> None:
        self.window.status_label.setText(text)

    def on_task_completed(self, summary_obj: object) -> None:
        if not isinstance(summary_obj, ConversionSummary):
            return

        summary = summary_obj
        self.state.last_summary = summary
        elapsed_str = f"{summary.elapsed_seconds:.1f}초" if summary.elapsed_seconds is not None else "알 수 없음"

        if summary.failed_count == 0 and summary.canceled_count == 0:
            self.window.toast.show_message(
                f"✅ 성공 {summary.success_count}개, 건너뜀 {summary.skipped_count}개 ({elapsed_str})",
                "🎉",
            )
        else:
            self.window.toast.show_message(
                f"⚠️ 성공 {summary.success_count} / 실패 {summary.failed_count} / 취소 {summary.canceled_count} ({elapsed_str})",
                "⚠️",
            )

        if summary.progid_used:
            self.window.hwp_status_label.setText("🟢 한글 연결됨")
        elif summary.failed_count:
            self.window.hwp_status_label.setText("🔴 한글 연결 오류")
        else:
            self.window.hwp_status_label.setText("🟢 한글 대기중")
        if self.state.close_after_worker:
            return
        dialog = self.window._create_result_dialog(summary)
        dialog.exec()

    def on_error_occurred(self, error_msg: str) -> None:
        self.window.toast.show_message("변환 중 오류 발생", "❌")
        self.window.hwp_status_label.setText("🔴 한글 연결 오류")
        QMessageBox.critical(self.window, "오류", f"변환 중 오류 발생:\n{error_msg}")

    def on_worker_finished(self) -> None:
        self.window._set_converting_state(False)
        self.window.progress_bar.setValue(0)
        self.window.progress_label.setText("0 / 0")
        self.window.status_label.setText("대기 중")
        self.window.hwp_status_label.setText("🟢 한글 대기중")

        if self.state.worker:
            try:
                self.state.worker.progress_updated.disconnect()
                self.state.worker.status_updated.disconnect()
                self.state.worker.task_completed.disconnect()
                self.state.worker.error_occurred.disconnect()
                self.state.worker.finished.disconnect()
            except (TypeError, RuntimeError):
                pass

        self.state.worker = None
        self.state.plan = None
        if self.state.close_after_worker:
            QTimer.singleShot(0, self.window.close)
