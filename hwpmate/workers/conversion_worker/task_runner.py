"""Qt 비의존 작업 실행기 — GUI ConversionWorker 와 헤드리스 CLI 가 공유한다.

단일 작업: 백업 → 런타임 출력 경로 할당 → 출력 폴더 준비 → 입력 확인 → 변환(재시도) → 감사 필드.
"""

from __future__ import annotations

import time
from dataclasses import dataclass
from pathlib import Path
from typing import Callable, Optional

from ...constants import MAX_RETRY_COUNT, RETRY_DELAY_SECONDS
from ...logging_config import get_logger
from ...models import ConversionTask, PlannedConversion
from ...services.hwp_print_settings import normalize_pdf_export_mode
from ...services.task_planner import TaskPlanner
from .backup import create_backup
from .protocol import ConverterEngine
from .summary import apply_converter_artifacts

logger = get_logger(__name__)

# 재순환 시 이전 한글 프로세스 종료 직후 Dispatch 가 실패할 수 있어 짧게 재시도한다.
RECYCLE_INITIALIZE_ATTEMPTS = 3
RECYCLE_INITIALIZE_DELAY_SECONDS = 1.5

StageCallback = Callable[[str], None]
StatusCallback = Callable[[str], None]


@dataclass(frozen=True)
class TaskRunOptions:
    format_type: str
    overwrite: bool = False
    backup_enabled: bool = True
    backup_max_files_per_stem: int = 20
    retry_count: int = 1
    pdf_export_mode: str = "saveas_first"

    @classmethod
    def from_plan(cls, plan: PlannedConversion) -> "TaskRunOptions":
        return cls(
            format_type=plan.format_type,
            overwrite=bool(plan.overwrite),
            backup_enabled=bool(plan.backup_enabled),
            backup_max_files_per_stem=int(getattr(plan, "backup_max_files_per_stem", 20) or 20),
            retry_count=max(0, min(MAX_RETRY_COUNT, int(plan.retry_count))),
            pdf_export_mode=normalize_pdf_export_mode(getattr(plan, "pdf_export_mode", None)),
        )


def _noop(_: str) -> None:
    return None


def execute_task(
    task: ConversionTask,
    converter: ConverterEngine,
    options: TaskRunOptions,
    *,
    planner: TaskPlanner,
    used_output_path_keys: set[str],
    cancel_check: Callable[[], bool],
    create_backup_fn: Optional[Callable[[Path], Path]] = None,
    on_stage: StageCallback = _noop,
    on_status: StatusCallback = _noop,
) -> list[str]:
    """작업 1건을 실행하고 task 상태를 갱신한다. 런타임 경고 목록을 반환."""
    runtime_warnings: list[str] = []
    backup = create_backup_fn or (
        lambda path: create_backup(path, max_files=options.backup_max_files_per_stem)
    )

    if options.backup_enabled:
        on_stage("백업 생성")
        try:
            task.backup_file = backup(task.input_file)
        except Exception as e:
            task.backup_error = str(e)
            logger.warning(f"백업 실패 (계속 진행): {e}")

    original_output = task.output_file
    on_stage("출력 경로 확인")
    if planner.allocate_output_path(
        task,
        used_path_keys=used_output_path_keys,
        overwrite=options.overwrite,
        format_type=options.format_type,
    ):
        warning = f"변환 직전 출력 충돌 감지로 경로 변경: {original_output} -> {task.output_file}"
        runtime_warnings.append(warning)
        logger.warning(warning)

    try:
        on_stage("출력 폴더 준비")
        task.output_file.parent.mkdir(parents=True, exist_ok=True)
    except Exception as e:
        task.status = "실패"
        task.error = f"폴더 생성 실패: {e}"
        return runtime_warnings

    if not task.input_file.exists():
        task.status = "실패"
        task.error = f"파일을 찾을 수 없음: {task.input_file.name}"
        logger.warning(f"파일 없음: {task.input_file}")
        return runtime_warnings

    task.status = "진행중"
    success = False
    error: str | None = None
    for attempt in range(options.retry_count + 1):
        if cancel_check():
            break

        on_stage("COM 내보내기")
        success, error = converter.convert_file(
            task.input_file,
            task.output_file,
            options.format_type,
            cancel_check=cancel_check,
        )
        if success:
            apply_converter_artifacts(task, converter)
            break

        if attempt < options.retry_count:
            on_stage("재시도 대기")
            task.retry_count += 1
            on_status(f"재시도 중: {task.input_file.name} ({task.retry_count}/{options.retry_count})")
            time.sleep(RETRY_DELAY_SECONDS)

    if success:
        task.status = "성공"
        task.error = None
    elif cancel_check():
        # 취소 요청 후 실패(또는 미완료)는 실패 대신 취소로 집계한다.
        detail = error.strip() if error else "사용자 취소"
        task.status = "취소됨"
        task.error = detail if detail == "사용자 취소" else f"사용자 취소 ({detail})"
    else:
        task.status = "실패"
        task.error = error
    return runtime_warnings


def recycle_converter(converter: ConverterEngine, options: TaskRunOptions) -> Optional[str]:
    """한글 프로세스를 정리 후 재초기화한다. 실패 시 오류 메시지를 반환."""
    try:
        converter.cleanup()
    except Exception as e:
        logger.warning(f"재순환 정리 중 오류(계속): {e}")

    last_error: Exception | None = None
    for attempt in range(RECYCLE_INITIALIZE_ATTEMPTS):
        try:
            converter.initialize(manage_com_apartment=False)
            if hasattr(converter, "pdf_export_mode"):
                converter.pdf_export_mode = options.pdf_export_mode
            return None
        except Exception as e:
            last_error = e
            logger.warning(
                f"한글 프로세스 재순환 초기화 실패 ({attempt + 1}/{RECYCLE_INITIALIZE_ATTEMPTS}): {e}"
            )
            if attempt + 1 < RECYCLE_INITIALIZE_ATTEMPTS:
                time.sleep(RECYCLE_INITIALIZE_DELAY_SECONDS)
    return str(last_error) if last_error is not None else "알 수 없는 오류"


def fail_pending_tasks(tasks: list[ConversionTask], message: str) -> int:
    """대기 중인 작업을 즉시 실패로 표시한다 (재순환 실패 등으로 더 진행할 수 없을 때)."""
    count = 0
    for task in tasks:
        if task.status == "대기":
            task.status = "실패"
            task.error = message
            count += 1
    return count


def compat_dialog_note(converter: object) -> Optional[str]:
    count = int(getattr(converter, "compat_dialog_responses", 0) or 0)
    if count <= 0:
        return None
    return (
        f"한글 「호환 문서(배치가 변경될 수 있습니다)」 확인 창 {count}회에 자동으로 「계속」했습니다. "
        "DOCX/RTF 등 호환 형식은 원본과 배치가 다를 수 있으니 결과를 확인해 주세요."
    )
