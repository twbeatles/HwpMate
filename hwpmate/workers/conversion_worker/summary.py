"""변환 결과 요약·엔진 상태·경고 수집."""

from __future__ import annotations

from pathlib import Path

from ...models import ConversionSummary, ConversionTask, PlannedConversion
from .protocol import ConverterEngine


def apply_converter_artifacts(task: ConversionTask, converter: ConverterEngine) -> None:
    created_files = getattr(converter, "last_created_files", [])
    task.created_files = [Path(path) for path in created_files]
    task.output_size = getattr(converter, "last_output_size", None)
    task.output_mtime = getattr(converter, "last_output_mtime", None)
    task.save_format = getattr(converter, "last_save_format", None)
    task.export_method = getattr(converter, "last_export_method", None)
    task.progid_used = converter.progid_used


def build_summary(
    *,
    format_type: str,
    tasks: list[ConversionTask],
    planned: PlannedConversion,
    warnings: list[str],
    elapsed_seconds: float,
    progid_used: str | None,
) -> ConversionSummary:
    """UI 스레드에 넘기기 전 작업 스냅샷을 복사한다."""
    task_snapshots = [task.snapshot() for task in tasks]
    skipped_snapshots = [task.snapshot() for task in planned.skipped_tasks]
    return ConversionSummary(
        format_type=format_type,
        tasks=task_snapshots + skipped_snapshots,
        warnings=list(warnings),
        elapsed_seconds=elapsed_seconds,
        progid_used=progid_used,
    )


def collect_converter_warnings(converter: ConverterEngine) -> list[str]:
    warnings: list[str] = []
    if getattr(converter, "security_module_registered", None) is False:
        detail = getattr(converter, "security_module_error", None)
        message = (
            "한글 보안 모듈 등록에 실패했습니다. "
            "파일 접근 시 '모두 허용' 창이 뜰 수 있으며, "
            "한컴 보안 모듈(FilePathCheckDLL) 등록이 근본 해결책입니다."
        )
        if detail:
            message += f" 상세: {detail}"
        warnings.append(message)

    process_warning = getattr(converter, "process_tracking_warning", None)
    if process_warning:
        warnings.append(str(process_warning))
    return warnings


def engine_status_payload(converter: ConverterEngine) -> dict:
    owned = getattr(converter, "owned_pids", None) or set()
    return {
        "security_module_registered": getattr(converter, "security_module_registered", None),
        "owned_pids": sorted(int(pid) for pid in owned),
        "snapshot_unreliable": bool(getattr(converter, "snapshot_unreliable", False)),
        "process_tracking_warning": getattr(converter, "process_tracking_warning", None),
        "progid_used": converter.progid_used,
    }
