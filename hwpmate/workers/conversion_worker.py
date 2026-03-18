from __future__ import annotations

import shutil
import time
from datetime import datetime
from pathlib import Path
from typing import Callable, Optional, Protocol

from PyQt6.QtCore import QThread, pyqtSignal

from ..logging_config import get_logger
from ..models import ConversionSummary, PlannedConversion
from ..services.hwp_converter import HWPConverter, pythoncom

logger = get_logger(__name__)


class ConverterEngine(Protocol):
    @property
    def progid_used(self) -> str | None: ...

    def initialize(self) -> bool: ...
    def convert_file(self, input_path, output_path, format_type="PDF") -> tuple[bool, str | None]: ...
    def cleanup(self) -> None: ...
    def has_owned_processes(self) -> bool: ...
    def kill_owned_processes(self) -> bool: ...


class ConversionWorker(QThread):
    """변환 작업 워커 스레드"""

    progress_updated = pyqtSignal(int, int, str)  # current, total, filename
    status_updated = pyqtSignal(str)
    task_completed = pyqtSignal(object)  # ConversionSummary
    error_occurred = pyqtSignal(str)

    _com_initialized = False

    def __init__(
        self,
        planned_conversion: PlannedConversion,
        converter_factory: Optional[Callable[[], ConverterEngine]] = None,
    ) -> None:
        super().__init__()
        self.planned_conversion = planned_conversion
        self.tasks = planned_conversion.tasks
        self.format_type = planned_conversion.format_type
        self.cancel_requested = False
        self.converter: Optional[ConverterEngine] = None
        self._converter_factory: Callable[[], ConverterEngine] = converter_factory or HWPConverter

    def cancel(self) -> None:
        """취소 요청"""
        self.cancel_requested = True

    def can_force_terminate(self) -> bool:
        converter = self.converter
        return bool(converter and converter.has_owned_processes())

    def run(self) -> None:
        """변환 작업 수행"""
        if pythoncom is not None:
            try:
                pythoncom.CoInitialize()
                self._com_initialized = True
            except Exception as e:
                logger.debug(f"Worker COM 초기화: {e}")

        start_ts = time.perf_counter()
        converter = self._converter_factory()
        self.converter = converter
        total = len(self.tasks)

        try:
            self.status_updated.emit("한글 프로그램 연결 중...")
            converter.initialize()
            self.status_updated.emit(f"연결 성공: {converter.progid_used}")

            for idx, task in enumerate(self.tasks):
                if self.cancel_requested:
                    self.status_updated.emit("사용자가 취소했습니다.")
                    break

                self.progress_updated.emit(idx, total, task.input_file.name)

                try:
                    self._create_backup(task.input_file)
                except Exception as e:
                    logger.warning(f"백업 실패 (계속 진행): {e}")

                try:
                    task.output_file.parent.mkdir(parents=True, exist_ok=True)
                except Exception as e:
                    task.status = "실패"
                    task.error = f"폴더 생성 실패: {e}"
                    continue

                if not task.input_file.exists():
                    task.status = "실패"
                    task.error = f"파일을 찾을 수 없음: {task.input_file.name}"
                    logger.warning(f"파일 없음: {task.input_file}")
                    continue

                task.status = "진행중"
                success, error = converter.convert_file(
                    task.input_file,
                    task.output_file,
                    self.format_type,
                )

                if success:
                    task.status = "성공"
                    task.error = None
                else:
                    task.status = "실패"
                    task.error = error

            if self.cancel_requested:
                for task in self.tasks:
                    if task.status == "대기":
                        task.status = "취소됨"
                        task.error = "사용자 취소"

            self.progress_updated.emit(total, total, "완료" if not self.cancel_requested else "취소됨")
            summary = ConversionSummary(
                format_type=self.format_type,
                tasks=list(self.tasks) + list(self.planned_conversion.skipped_tasks),
                warnings=list(self.planned_conversion.warnings),
                elapsed_seconds=time.perf_counter() - start_ts,
                progid_used=converter.progid_used,
            )
            self.task_completed.emit(summary)
        except Exception as e:
            logger.exception("변환 중 오류 발생")
            self.error_occurred.emit(str(e))
        finally:
            try:
                converter.cleanup()
            except Exception as e:
                logger.error(f"정리 중 오류: {e}")

            if self._com_initialized:
                if pythoncom is not None:
                    try:
                        pythoncom.CoUninitialize()
                    except Exception:
                        pass

    def force_terminate(self) -> bool:
        """앱이 소유한 한글 프로세스만 강제 종료."""
        converter = self.converter
        if converter is None:
            return False
        return converter.kill_owned_processes()

    def _create_backup(self, file_path: Path) -> None:
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
        except Exception as e:
            logger.error(f"백업 생성 중 오류: {e}")
            raise
