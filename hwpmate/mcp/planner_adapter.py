"""읽기 전용 변환 미리보기 (구현 설계서 §4.2).

- 실제 문서 변환·파일 생성을 하지 않는다 (출력 폴더도 만들지 않음).
- 출력 경로 할당은 기존 TaskPlanner.resolve_output_conflicts()를 재사용한다.
- plan_id는 비추측형 핸들이며 입력 stat과 묶어 PLAN_STALE을 검출한다.
"""

from __future__ import annotations

import os
import secrets
import threading
from dataclasses import dataclass, field
from datetime import datetime, timedelta, timezone
from pathlib import Path

from ..constants import BACKUP_DIR_NAME, FORMAT_TYPES, SUPPORTED_EXTENSIONS
from ..path_utils import is_path_length_blocking, is_path_length_risky
from ..services.task_planner import TaskPlanner
from ..models import ConversionTask
from . import MCP_SCHEMA_VERSION
from .config import McpConfig
from .errors import McpError
from .policy import (
    check_input_path,
    check_output_dir,
    is_within_directory,
    normalize_format,
)

_PREVIEW_ITEM_LIMIT = 100

_plans: dict[str, "PlanRecord"] = {}
_plans_lock = threading.Lock()


@dataclass
class PlanEntry:
    input_file: str
    size: int
    mtime_ns: int


@dataclass
class PlanRecord:
    plan_id: str
    format: str
    output_dir: str
    recursive: bool
    entries: list[PlanEntry] = field(default_factory=list)
    expires_at: datetime = field(
        default_factory=lambda: datetime.now(timezone.utc) + timedelta(seconds=600)
    )

    def is_expired(self, now: datetime | None = None) -> bool:
        return (now or datetime.now(timezone.utc)) >= self.expires_at


def _stat_entry(path: Path) -> PlanEntry:
    try:
        stat = path.stat()
    except OSError as exc:
        raise McpError("INPUT_NOT_FOUND", f"입력 파일을 읽을 수 없습니다: {path} ({exc})") from exc
    return PlanEntry(input_file=str(path), size=stat.st_size, mtime_ns=stat.st_mtime_ns)


def _iter_scan_files(
    root: Path,
    *,
    recursive: bool,
    max_depth: int,
    output_dir: Path | None,
) -> tuple[list[Path], int]:
    """폴더에서 변환 후보를 수집한다. backup·숨김 폴더·출력 폴더를 제외한다."""
    allowed = {ext.lower() for ext in SUPPORTED_EXTENSIONS}
    excluded = BACKUP_DIR_NAME.lower()
    found: list[Path] = []
    skipped_backup = 0
    base_depth = len(root.parts)
    if recursive:
        for dirpath, dirnames, filenames in os.walk(root):
            try:
                rel_depth = len(Path(dirpath).parts) - base_depth
            except OSError:
                continue
            if rel_depth > max_depth:
                dirnames[:] = []
                continue
            current = Path(dirpath)
            if output_dir is not None and (
                current == output_dir or is_within_directory(current, output_dir)
            ):
                dirnames[:] = []
                continue
            dirnames[:] = [
                name
                for name in dirnames
                if name.lower() != excluded and not name.startswith(".")
            ]
            for filename in filenames:
                candidate = current / filename
                if output_dir is not None and is_within_directory(candidate, output_dir):
                    continue
                if candidate.suffix.lower() in allowed:
                    found.append(candidate)
    else:
        try:
            with os.scandir(root) as entries:
                for entry in entries:
                    try:
                        if not entry.is_file(follow_symlinks=False):
                            continue
                    except OSError:
                        continue
                    candidate = Path(entry.path)
                    if output_dir is not None and is_within_directory(candidate, output_dir):
                        continue
                    if candidate.suffix.lower() in allowed:
                        found.append(candidate)
        except OSError as exc:
            raise McpError("INPUT_NOT_FOUND", f"폴더를 읽을 수 없습니다: {root} ({exc})") from exc
    # backup 폴더 자체를 입력으로 준 경우 os.walk가 수집하므로 제외 집계
    kept: list[Path] = []
    for candidate in found:
        if excluded in {part.lower() for part in candidate.relative_to(root).parts[:-1]}:
            skipped_backup += 1
            continue
        kept.append(candidate)
    return sorted(kept, key=lambda p: str(p).lower()), skipped_backup


def _collect_inputs(
    cfg: McpConfig,
    raw_inputs: list[str],
    *,
    recursive: bool,
    output_dir: Path | None,
) -> tuple[list[Path], int]:
    if not raw_inputs:
        raise McpError("INPUT_DENIED", "입력 경로가 비어 있습니다.")
    ordered: dict[str, Path] = {}
    skipped_backup = 0
    for raw in raw_inputs:
        resolved = check_input_path(cfg, raw)
        if resolved.is_file():
            if resolved.suffix.lower() not in {ext.lower() for ext in SUPPORTED_EXTENSIONS}:
                raise McpError(
                    "BAD_INPUT",
                    f"한글 문서(.hwp/.hwpx)만 변환할 수 있습니다: {resolved.name}",
                )
            if output_dir is not None and is_within_directory(resolved, output_dir):
                continue
            ordered[str(resolved).lower()] = resolved
        elif resolved.is_dir():
            files, skipped = _iter_scan_files(
                resolved,
                recursive=recursive,
                max_depth=max(0, cfg.max_recursive_depth),
                output_dir=output_dir,
            )
            skipped_backup += skipped
            for candidate in files:
                ordered[str(candidate).lower()] = candidate
        else:
            raise McpError("INPUT_NOT_FOUND", f"입력 경로가 존재하지 않습니다: {resolved}")
    return sorted(ordered.values(), key=lambda p: str(p).lower()), skipped_backup


def _check_quotas(cfg: McpConfig, files: list[Path]) -> int:
    if len(files) > max(1, cfg.max_files_per_job):
        raise McpError(
            "INPUT_TOO_LARGE",
            f"입력 파일이 너무 많습니다: {len(files)}개 (최대 {cfg.max_files_per_job}개)",
        )
    total = 0
    for candidate in files:
        try:
            total += candidate.stat().st_size
        except OSError as exc:
            raise McpError("INPUT_NOT_FOUND", f"입력 파일을 읽을 수 없습니다: {candidate}") from exc
    limit = max(1, cfg.max_input_total_mb) * 1024 * 1024
    if total > limit:
        raise McpError(
            "INPUT_TOO_LARGE",
            f"입력 총 크기가 제한을 초과했습니다: {total} bytes (최대 {cfg.max_input_total_mb}MB)",
        )
    return total


def preview_conversion(
    cfg: McpConfig,
    *,
    input_paths: list[str],
    format_type: str,
    output_dir: str,
    recursive: bool = False,
) -> dict:
    clear_expired_plans()
    fmt = normalize_format(format_type)
    if not str(output_dir or "").strip():
        raise McpError("OUTPUT_DENIED", "MCP 변환은 출력 폴더를 명시해야 합니다 (원본 근처 자동 저장 금지).")
    approved_output = check_output_dir(cfg, output_dir)
    files, skipped_backup = _collect_inputs(
        cfg, [str(path) for path in input_paths], recursive=bool(recursive), output_dir=approved_output
    )
    if not files:
        raise McpError("INPUT_NOT_FOUND", "변환할 한글 문서(.hwp/.hwpx)가 없습니다.")
    _check_quotas(cfg, files)

    output_ext = str(FORMAT_TYPES[fmt]["ext"])
    tasks: list[ConversionTask] = []
    skipped = 0
    for input_file in files:
        if input_file.suffix.lower() == output_ext.lower():
            skipped += 1
            continue
        tasks.append(
            ConversionTask(
                input_file=input_file,
                output_file=approved_output / (input_file.stem + output_ext),
            )
        )
    planner = TaskPlanner()
    conflicts_renamed = planner.resolve_output_conflicts(tasks, overwrite=False, format_type=fmt)

    warnings: list[str] = []
    if skipped_backup:
        warnings.append(f"백업(backup) 폴더의 {skipped_backup}개 파일은 수집에서 제외했습니다.")
    for task in tasks:
        if is_path_length_blocking(task.output_file) or is_path_length_blocking(task.input_file):
            warnings.append(f"경로가 너무 길어 COM 변환이 실패할 수 있습니다: {task.input_file.name}")
            break
    else:
        for task in tasks:
            if is_path_length_risky(task.output_file) or is_path_length_risky(task.input_file):
                warnings.append("경로 길이가 240자를 넘는 파일이 있어 주의가 필요합니다.")
                break
    if not approved_output.exists():
        warnings.append("출력 폴더가 아직 없어 제출 시 생성됩니다.")
    warnings.append("변환 계획 시점과 실행 시점 사이 파일명이 바뀔 수 있습니다")

    entries = [_stat_entry(path) for path in files]
    plan_id = secrets.token_urlsafe(24)
    record = PlanRecord(
        plan_id=plan_id,
        format=fmt,
        output_dir=str(approved_output),
        recursive=bool(recursive),
        entries=entries,
        expires_at=datetime.now(timezone.utc) + timedelta(seconds=max(60, cfg.plan_ttl_seconds)),
    )
    with _plans_lock:
        _plans[plan_id] = record
    preview_items = [
        {"input": str(task.input_file), "output": str(task.output_file)}
        for task in tasks[:_PREVIEW_ITEM_LIMIT]
    ]
    return {
        "schema_version": MCP_SCHEMA_VERSION,
        "ok": True,
        "plan_id": plan_id,
        "expires_at": record.expires_at.isoformat(),
        "format": fmt,
        "requested": len(files),
        "planned": len(tasks),
        "skipped": skipped,
        "conflicts_renamed": conflicts_renamed,
        "requires_confirmation": True,
        "warnings": warnings,
        "preview": preview_items,
    }


def get_plan(plan_id: str) -> PlanRecord:
    now = datetime.now(timezone.utc)
    with _plans_lock:
        record = _plans.get(str(plan_id))
        if record is None or record.is_expired(now):
            _plans.pop(str(plan_id), None)
            raise McpError("PLAN_EXPIRED", "계획이 만료되었거나 존재하지 않습니다. 다시 preview를 요청하세요.")
    try:
        current = [
            (entry.input_file, Path(entry.input_file).stat().st_size, Path(entry.input_file).stat().st_mtime_ns)
            for entry in record.entries
        ]
    except OSError as exc:
        raise McpError("PLAN_STALE", f"입력 파일이 변경·삭제되어 계획을 사용할 수 없습니다: {exc}") from exc
    for entry, (path, size, mtime_ns) in zip(record.entries, current):
        if entry.input_file != path or entry.size != size or entry.mtime_ns != mtime_ns:
            raise McpError("PLAN_STALE", "입력 파일이 변경되어 계획을 사용할 수 없습니다. 다시 preview를 요청하세요.")
    return record


def clear_expired_plans() -> int:
    now = datetime.now(timezone.utc)
    with _plans_lock:
        expired = [key for key, record in _plans.items() if record.is_expired(now)]
        for key in expired:
            del _plans[key]
    return len(expired)
