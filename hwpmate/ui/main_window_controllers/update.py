from __future__ import annotations

import os
import sys
from pathlib import Path
from typing import TYPE_CHECKING, Optional

from PyQt6.QtCore import QObject, QThread, QTimer, pyqtSignal

from ...constants import (
    UPDATE_MANIFEST_URL_DEFAULT,
    UPDATE_PUBLIC_KEY_B64_DEFAULT,
    VERSION,
)
from ...logging_config import get_logger
from ...services.update_installer import (
    cleanup_update_staging,
    consume_update_result,
    launch_update_helper,
    prepare_staged_update,
    resolve_update_staging_root,
    stream_update_artifact,
    unique_update_backup_path,
    update_result_path,
)
from ...services.update_manifest import (
    NoUpdateAvailableError,
    ReleaseManifest,
    download_release_manifest,
    verify_release_manifest,
)
from ..dialogs.update_dialog import UpdateDialog

if TYPE_CHECKING:
    from ..main_window import MainWindow

logger = get_logger(__name__)


class UpdateCheckWorker(QThread):
    """업데이트 매니페스트 확인 워커."""

    finished_with_update = pyqtSignal(object)  # ReleaseManifest
    finished_no_update = pyqtSignal()
    finished_with_error = pyqtSignal(str)

    def __init__(self, manifest_url: str, public_key: str, current_version: str) -> None:
        super().__init__()
        self.manifest_url = manifest_url
        self.public_key = public_key
        self.current_version = current_version

    def run(self) -> None:
        try:
            raw_manifest = download_release_manifest(self.manifest_url)
            manifest = verify_release_manifest(
                raw_manifest,
                public_key=self.public_key,
                current_version=self.current_version,
            )
            self.finished_with_update.emit(manifest)
        except NoUpdateAvailableError:
            self.finished_no_update.emit()
        except Exception as exc:
            logger.warning(f"업데이트 확인 실패: {exc}")
            self.finished_with_error.emit(str(exc))


class UpdateDownloadWorker(QThread):
    """업데이트 파일 다운로드 워커."""

    progress = pyqtSignal(int, int)  # current, total
    download_finished = pyqtSignal(object)  # Path (staged file)
    download_failed = pyqtSignal(str)

    def __init__(self, manifest: ReleaseManifest, staging_root: Path) -> None:
        super().__init__()
        self.manifest = manifest
        self.staging_root = staging_root
        self._is_cancelled = False

    def cancel(self) -> None:
        """다운로드 취소 요청."""
        self._is_cancelled = True

    def run(self) -> None:
        try:
            chunks = stream_update_artifact(
                self.manifest,
                cancel_check=lambda: self._is_cancelled,
            )
            staged = prepare_staged_update(
                self.manifest,
                chunks=chunks,
                staging_root=self.staging_root,
                progress_callback=self._on_progress,
                cancel_check=lambda: self._is_cancelled,
            )
            if self._is_cancelled:
                logger.info("업데이트 다운로드 취소됨")
                self.download_failed.emit("사용자에 의해 다운로드가 취소되었습니다.")
            elif staged is not None:
                self.download_finished.emit(staged)
        except Exception as exc:
            if not self._is_cancelled:
                logger.error(f"업데이트 다운로드 실패: {exc}")
                self.download_failed.emit(str(exc))


    def _on_progress(self, current: int, total: int) -> None:
        self.progress.emit(current, total)


class UpdateController(QObject):
    """메인 윈도우용 업데이트 컨트롤러."""

    def __init__(self, window: MainWindow) -> None:
        super().__init__(window)
        self.window = window
        self.staging_root = resolve_update_staging_root()
        self._check_worker: Optional[UpdateCheckWorker] = None
        self._download_worker: Optional[UpdateDownloadWorker] = None
        self._staged_file: Optional[Path] = None
        # 스테이징 파일을 받을 때 사용한 매니페스트 (이후 새 버전 매니페스트와 혼동 방지)
        self._staged_manifest: Optional[ReleaseManifest] = None
        self._current_manifest: Optional[ReleaseManifest] = None
        self._active_dialog: Optional[UpdateDialog] = None
        self._is_manual_check = False
        # showEvent 는 트레이/최소화 복원 때도 다시 호출되므로 시작 작업은 1회만 수행한다.
        self._previous_result_checked = False
        self._startup_check_scheduled = False

    def check_previous_update_result(self) -> None:
        """앱 시작 시 이전 업데이트 실행 결과를 확인하고 피드백을 제공합니다."""
        if self._previous_result_checked:
            return
        self._previous_result_checked = True
        res_path = update_result_path(self.staging_root)
        result = consume_update_result(res_path)
        # 이전 업데이트의 헬퍼 exe·미적용 스테이징 파일 정리 (실행 중이면 조용히 건너뜀)
        try:
            removed = cleanup_update_staging(self.staging_root)
            if removed:
                logger.info(f"업데이트 스테이징 잔여 파일 정리: {removed}개")
        except Exception as exc:
            logger.debug(f"업데이트 스테이징 정리 실패(무시): {exc}")
        if not result:
            return

        status = str(result.get("status", ""))
        if status == "applied":
            self.window.toast.show_message(
                f"HwpMate v{VERSION} 업데이트가 성공적으로 적용되었습니다! 🎉",
                icon="🎉",
                duration=4000,
            )
            logger.info("이전 업데이트 적용 성공 확인")
        elif status == "rolled_back":
            err = str(result.get("error", "알 수 없는 오류"))
            self.window.toast.show_message(
                f"업데이트 적용 실패로 이전 버전으로 복구되었습니다: {err}",
                icon="⚠️",
                duration=6000,
            )
            logger.warning(f"이전 업데이트 롤백 확인: {err}")
        elif status == "failed":
            err = str(result.get("error", "알 수 없는 오류"))
            self.window.toast.show_message(
                f"업데이트 적용 실패: {err}",
                icon="❌",
                duration=6000,
            )
            logger.error(f"이전 업데이트 실패 확인: {err}")

    def schedule_startup_check(self, delay_ms: int = 3000) -> None:
        """시작 후 일정 시간 뒤 조용히 업데이트를 확인합니다 (프로세스당 1회)."""
        if self._startup_check_scheduled:
            return
        self._startup_check_scheduled = True
        QTimer.singleShot(delay_ms, lambda: self.start_check(manual=False))

    def start_check(self, *, manual: bool = False) -> None:
        """업데이트 확인을 시작합니다."""
        if self._check_worker and self._check_worker.isRunning():
            if manual:
                self.window.toast.show_message("업데이트를 확인하는 중입니다...", icon="ℹ️")
            return

        self._is_manual_check = manual
        if manual:
            self.window.toast.show_message("최신 버전을 확인하고 있습니다...", icon="ℹ️")

        self._check_worker = UpdateCheckWorker(
            manifest_url=UPDATE_MANIFEST_URL_DEFAULT,
            public_key=UPDATE_PUBLIC_KEY_B64_DEFAULT,
            current_version=VERSION,
        )
        self._check_worker.finished_with_update.connect(self._on_update_found)
        self._check_worker.finished_no_update.connect(self._on_no_update)
        self._check_worker.finished_with_error.connect(self._on_check_error)
        self._check_worker.start()

    def _on_update_found(self, manifest: ReleaseManifest) -> None:
        self._current_manifest = manifest
        logger.info(f"새 업데이트 발견: v{manifest.version}")
        self._show_update_dialog(manifest)

    def _on_no_update(self) -> None:
        logger.info("현재 최신 버전을 사용 중입니다.")
        if self._is_manual_check:
            self.window.toast.show_message(
                f"현재 최신 버전(v{VERSION})을 사용 중입니다.",
                icon="✅",
                duration=3000,
            )

    def _on_check_error(self, error: str) -> None:
        logger.warning(f"업데이트 확인 실패: {error}")
        if self._is_manual_check:
            self.window.toast.show_message(
                f"업데이트 확인에 실패했습니다.\n네트워크 상태를 확인해 주세요.",
                icon="⚠️",
                duration=4000,
            )

    def _show_update_dialog(self, manifest: ReleaseManifest) -> None:
        if self._active_dialog and self._active_dialog.isVisible():
            self._active_dialog.raise_()
            self._active_dialog.activateWindow()
            return

        is_dark = getattr(self.window, "current_theme", "dark") == "dark"
        dialog = UpdateDialog(self.window, manifest, is_dark=is_dark)
        self._active_dialog = dialog
        dialog.update_accepted.connect(lambda: self._start_download(manifest, dialog))
        dialog.download_cancelled.connect(self._cancel_download)
        dialog.apply_restart_requested.connect(self._apply_and_restart)
        if self._has_staged_file_for(manifest):
            # 이미 검증까지 끝난 파일이 있으면 다시 받지 않는다.
            dialog.set_download_complete()
        dialog.show()

    def _has_staged_file_for(self, manifest: ReleaseManifest) -> bool:
        staged = self._staged_file
        staged_manifest = self._staged_manifest
        if staged is None or staged_manifest is None:
            return False
        if (
            staged_manifest.version != manifest.version
            or staged_manifest.artifact_sha256 != manifest.artifact_sha256
        ):
            return False
        try:
            return staged.is_file() and staged.stat().st_size == manifest.artifact_size
        except OSError:
            return False

    def _busy_reason(self) -> Optional[str]:
        state = getattr(self.window, "state", None)
        conversion = getattr(self.window, "conversion_controller", None)
        if conversion is not None and conversion.is_conversion_active():
            return "변환 작업이 진행 중입니다"
        if state is not None and getattr(state, "is_planning", False):
            return "변환 준비가 진행 중입니다"
        return None

    def _cancel_download(self) -> None:
        if self._download_worker and self._download_worker.isRunning():
            logger.info("다운로드 취소 요청 수신")
            self._download_worker.cancel()

    def _start_download(self, manifest: ReleaseManifest, dialog: UpdateDialog) -> None:
        if self._download_worker and self._download_worker.isRunning():
            return

        dialog.set_downloading()
        self._download_worker = UpdateDownloadWorker(manifest, self.staging_root)
        self._download_worker.progress.connect(dialog.update_progress)
        self._download_worker.download_finished.connect(
            lambda staged: self._on_download_finished(staged, dialog, manifest)
        )
        self._download_worker.download_failed.connect(dialog.set_download_failed)
        self._download_worker.start()


    def _on_download_finished(
        self,
        staged: Path,
        dialog: UpdateDialog,
        manifest: Optional[ReleaseManifest] = None,
    ) -> None:
        self._staged_file = staged
        self._staged_manifest = manifest or self._current_manifest
        logger.info(f"업데이트 파일 다운로드 완료: {staged}")
        dialog.set_download_complete()

    def _apply_and_restart(self) -> None:
        staged_manifest = self._staged_manifest or self._current_manifest
        if not self._staged_file or not staged_manifest:
            logger.error("적용할 스테이징 파일 또는 매니페스트 정보가 없습니다.")
            return

        busy = self._busy_reason()
        if busy:
            # 변환 중 종료하면 결과 리포트 없이 중단되고 한글 프로세스가 고아로 남을 수 있다.
            self.window.toast.show_message(
                f"{busy}. 작업이 끝난 뒤 「🔄 업데이트」에서 다시 적용해 주세요. "
                "(다운로드한 파일은 유지됩니다)",
                icon="⚠️",
                duration=5000,
            )
            logger.info(f"업데이트 적용 보류: {busy}")
            return

        target = Path(sys.executable).resolve()
        if not bool(getattr(sys, "frozen", False)):
            self.window.toast.show_message(
                "개발 환경(파이썬 스크립트 실행)에서는 자동 교체가 건너뛰어집니다.",
                icon="ℹ️",
                duration=4000,
            )
            logger.info("개발 환경에서 재시작 요청 무시")
            return

        backup = unique_update_backup_path(target, VERSION)
        res_file = update_result_path(self.staging_root)

        try:
            launch_update_helper(
                target=target,
                staged=self._staged_file,
                backup=backup,
                parent_pid=os.getpid(),
                expected_sha256=staged_manifest.artifact_sha256,
                expected_size=staged_manifest.artifact_size,
                result_file=res_file,
            )
        except Exception as exc:
            logger.error(f"업데이트 헬퍼 실행 실패: {exc}")
            self.window.toast.show_message(
                f"업데이트 프로세스 실행 실패: {exc}",
                icon="❌",
            )
            return

        logger.info("업데이트 헬퍼 프로세스 시작 완료, 애플리케이션 종료")
        from PyQt6.QtWidgets import QApplication

        # 헬퍼는 이 프로세스 종료를 실제로 기다린다. 창 종료 흐름(설정 저장·스캔 정리)을 거친 뒤 끝낸다.
        self.window.close()
        app = QApplication.instance()
        if app:
            app.quit()
