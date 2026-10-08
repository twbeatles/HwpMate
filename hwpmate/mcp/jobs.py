"""단일 worker job 큐: idempotency·recovery·상태 조회 (구현 설계서 §4.4·§5.2).

- 동시 변환 최대 1개. 기존 GUI/CLI 잠금을 우회·삭제하지 않는다.
- 외부 잠금 경쟁은 CLI 실패(stderr "이미 실행 중")로 감지해 rejected_busy로 보고한다.
- MCP 재시작으로 날아간 running/queued 작업은 interrupted로 복원하며 자동 재실행하지 않는다.
"""

from __future__ import annotations

import hashlib
import json
import os
import secrets
import threading
import time
from dataclasses import asdict, dataclass, field
from datetime import datetime, timezone
from pathlib import Path

from . import MCP_SCHEMA_VERSION
from .config import McpConfig
from .errors import McpError
from .logging_utils import setup_stderr_logging
from .planner_adapter import PlanRecord
from .report_reader import classify_job, read_cli_report
from .runner import (
    build_cli_argv,
    is_busy_failure,
    is_com_unavailable_failure,
    resolve_cli_executable,
    run_cli,
)
from .schemas import CONFIRMATION_TOKEN

logger = setup_stderr_logging()

TERMINAL_STATES = frozenset(
    {
        "succeeded",
        "partially_failed",
        "failed",
        "canceled",
        "timeout",
        "interrupted",
        "rejected_busy",
        "expired",
    }
)


def check_platform_supported() -> None:
    if os.name != "nt":
        raise McpError("PLATFORM_UNSUPPORTED", "HwpMate 변환은 Windows + 한컴오피스 한글 환경에서만 지원됩니다.")


def recheck_output(cfg: McpConfig, record: JobRecord) -> Path:
    """실행 직전 출력 폴더를 허용 정책으로 재검증한다 (TOCTOU 대응)."""
    from .policy import check_output_dir

    return check_output_dir(cfg, record.output_dir)


def _utcnow_iso() -> str:
    return datetime.now(timezone.utc).isoformat()


def plan_hash(record: PlanRecord) -> str:
    digest = hashlib.sha256()
    digest.update(record.format.encode("utf-8"))
    digest.update(b"\x00" + record.output_dir.encode("utf-8"))
    for entry in record.entries:
        digest.update(
            f"\x00{entry.input_file}\x00{entry.size}\x00{entry.mtime_ns}".encode("utf-8")
        )
    return digest.hexdigest()


@dataclass
class JobRecord:
    job_id: str
    format: str = ""
    output_dir: str = ""
    inputs: list[str] = field(default_factory=list)
    idempotency_key: str = ""
    plan_hash: str = ""
    state: str = "queued"
    created_at: str = ""
    started_at: str = ""
    finished_at: str = ""
    success_count: int = 0
    failed_count: int = 0
    skipped_count: int = 0
    canceled_count: int = 0
    items: list[dict] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    error: dict = field(default_factory=dict)
    timed_out: bool = False

    def to_status_dict(self) -> dict:
        return {
            "schema_version": MCP_SCHEMA_VERSION,
            "ok": True,
            "job_id": self.job_id,
            "state": self.state,
            "format": self.format,
            "requested": len(self.inputs),
            "success_count": self.success_count,
            "failed_count": self.failed_count,
            "skipped_count": self.skipped_count,
            "canceled_count": self.canceled_count,
            "created_at": self.created_at,
            "started_at": self.started_at,
            "finished_at": self.finished_at,
            "error": dict(self.error),
        }


class JobManager:
    def __init__(self, cfg: McpConfig) -> None:
        self._cfg = cfg
        if cfg.max_parallel_conversions != 1:
            logger.warning(
                "max_parallel_conversions=%s 요청됐으나 단일 worker로 고정합니다.",
                cfg.max_parallel_conversions,
            )
        self._job_dir = cfg.effective_job_dir()
        self._job_dir.mkdir(parents=True, exist_ok=True)
        self._lock = threading.RLock()
        self._active: str | None = None
        self._queue: list[str] = []
        self._records: dict[str, JobRecord] = {}
        self._idem: dict[str, str] = {}
        self._load_all()
        self._cleanup_retention()
        self._recover_interrupted()

    # -- persistence -----------------------------------------------------
    def _record_path(self, job_id: str) -> Path:
        return self._job_dir / f"{job_id}.json"

    def _persist(self, record: JobRecord) -> None:
        with self._lock:
            tmp = self._record_path(record.job_id).with_suffix(".tmp")
            tmp.write_text(json.dumps(asdict(record), ensure_ascii=False, indent=2), encoding="utf-8")
            tmp.replace(self._record_path(record.job_id))
        if record.idempotency_key:
            self._idem[f"{record.idempotency_key}\x00{record.plan_hash}"] = record.job_id
            (self._job_dir / "idempotency.json").write_text(
                json.dumps(self._idem, ensure_ascii=False, indent=2), encoding="utf-8"
            )

    def _load_all(self) -> None:
        idem_path = self._job_dir / "idempotency.json"
        if idem_path.is_file():
            try:
                self._idem = json.loads(idem_path.read_text(encoding="utf-8"))
            except (ValueError, OSError):
                self._idem = {}
        for path in sorted(self._job_dir.glob("*.json")):
            if path.name == "idempotency.json":
                continue
            try:
                data = json.loads(path.read_text(encoding="utf-8"))
                record = JobRecord(**{k: data.get(k, v) for k, v in asdict(JobRecord(job_id="")).items()})
                record.job_id = path.stem
                self._records[record.job_id] = record
            except (ValueError, OSError, TypeError):
                continue

    def _cleanup_retention(self) -> None:
        days = self._cfg.job_retention_days
        if days <= 0:
            return
        cutoff = time.time() - days * 86400
        removed = 0
        for path in sorted(self._job_dir.glob("*.tmp")):
            try:
                path.unlink()
            except OSError:
                pass
        for job_id, record in list(self._records.items()):
            if record.state not in TERMINAL_STATES:
                continue
            stamp = record.finished_at or record.created_at
            try:
                age_ok = datetime.fromisoformat(stamp).timestamp() < cutoff
            except (ValueError, TypeError):
                try:
                    age_ok = self._record_path(job_id).stat().st_mtime < cutoff
                except OSError:
                    continue
            if not age_ok:
                continue
            try:
                self._record_path(job_id).unlink()
            except OSError:
                continue
            del self._records[job_id]
            removed += 1
        if removed:
            self._idem = {
                key: job_id for key, job_id in self._idem.items() if job_id in self._records
            }
            try:
                (self._job_dir / "idempotency.json").write_text(
                    json.dumps(self._idem, ensure_ascii=False, indent=2), encoding="utf-8"
                )
            except OSError:
                pass
            logger.info("보존 기간 만료 job 기록 %d건을 정리했습니다.", removed)

    def _recover_interrupted(self) -> None:
        changed = False
        for record in self._records.values():
            if record.state in ("queued", "running"):
                record.state = "interrupted"
                record.finished_at = record.finished_at or _utcnow_iso()
                record.error = {
                    "code": "INTERRUPTED",
                    "message": "MCP 서버 재시작으로 중단되었습니다. 파일 성공 여부가 불명확하므로 자동 재실행하지 않습니다.",
                }
                self._persist(record)
                changed = True
        if changed:
            logger.warning("중단된 job을 interrupted로 복원했습니다 (자동 재실행 없음).")

    # -- public API ------------------------------------------------------
    def submit(
        self, plan: PlanRecord, *, confirmation: str, idempotency_key: str
    ) -> JobRecord:
        if self._cfg.require_confirmation and confirmation != CONFIRMATION_TOKEN:
            raise McpError("CONFIRMATION_REQUIRED", "명시적 승인 토큰이 필요합니다.")
        key = str(idempotency_key or "").strip()
        if not key:
            raise McpError("CONFIRMATION_REQUIRED", "idempotency_key가 필요합니다.")
        digest = plan_hash(plan)
        with self._lock:
            existing_id = self._idem.get(f"{key}\x00{digest}")
            if existing_id and existing_id in self._records:
                return self._records[existing_id]
            for job_id, other in self._records.items():
                if other.idempotency_key == key and other.plan_hash != digest:
                    raise McpError(
                        "IDEMPOTENCY_CONFLICT",
                        f"같은 idempotency_key로 다른 계획이 이미 제출되었습니다: {job_id}",
                    )
            if self._active is not None and len(self._queue) >= max(0, self._cfg.max_queued_jobs):
                raise McpError("JOB_QUEUE_FULL", "대기 큐가 가득 찼습니다. 나중에 다시 요청하세요.")
            record = JobRecord(
                job_id=secrets.token_hex(16),
                format=plan.format,
                output_dir=plan.output_dir,
                inputs=[entry.input_file for entry in plan.entries],
                idempotency_key=key,
                plan_hash=digest,
                state="queued",
                created_at=_utcnow_iso(),
            )
            self._records[record.job_id] = record
            self._persist(record)
            if self._active is None:
                self._start_locked(record)
            else:
                self._queue.append(record.job_id)
                self._persist(record)
            return record

    def get(self, job_id: str) -> JobRecord:
        record = self._records.get(str(job_id))
        if record is None:
            raise McpError("JOB_NOT_FOUND", f"작업을 찾을 수 없습니다: {job_id}")
        return record

    def cancel(self, job_id: str, *, confirmation: str) -> JobRecord:
        if confirmation != CONFIRMATION_TOKEN:
            raise McpError("CONFIRMATION_REQUIRED", "취소에도 명시적 승인 토큰이 필요합니다.")
        with self._lock:
            record = self.get(job_id)
            if record.state in TERMINAL_STATES:
                return record
            if record.job_id in self._queue:
                self._queue.remove(record.job_id)
                record.state = "canceled"
                record.finished_at = _utcnow_iso()
                self._persist(record)
                return record
        raise McpError(
            "CANCELLATION_UNSUPPORTED_WHILE_RUNNING",
            "실행 중인 작업은 안전하게 취소할 수 없어 강제 종료하지 않습니다.",
        )

    # -- worker ----------------------------------------------------------
    def _start_locked(self, record: JobRecord) -> None:
        self._active = record.job_id
        thread = threading.Thread(target=self._run_job, args=(record.job_id,), daemon=True)
        thread.start()

    def _queue_expired(self, record: JobRecord) -> bool:
        limit = self._cfg.queue_wait_timeout_seconds
        if limit <= 0 or not record.created_at:
            return False
        try:
            waited = time.time() - datetime.fromisoformat(record.created_at).timestamp()
        except (ValueError, TypeError):
            return False
        return waited > limit

    def _run_job(self, job_id: str) -> None:
        try:
            record = self._records[job_id]
            if self._queue_expired(record):
                record.state = "expired"
                record.finished_at = _utcnow_iso()
                record.error = {
                    "code": "TIMEOUT",
                    "message": "대기 큐에서 만료되어 실행하지 않았습니다. 다시 preview 후 제출하세요.",
                }
                self._persist(record)
                return
            record.state = "running"
            record.started_at = _utcnow_iso()
            self._persist(record)
            self._execute(record)
        except McpError as exc:
            record = self._records[job_id]
            record.state = "failed"
            record.finished_at = _utcnow_iso()
            record.error = {"code": exc.code, "message": exc.message}
            self._persist(record)
        except Exception as exc:  # pragma: no cover - 방어적
            record = self._records[job_id]
            record.state = "failed"
            record.finished_at = _utcnow_iso()
            record.error = {"code": "INTERNAL", "message": str(exc)}
            self._persist(record)
        finally:
            with self._lock:
                if self._active == job_id:
                    self._active = None
                if self._queue:
                    next_id = self._queue.pop(0)
                    self._start_locked(self._records[next_id])

    def _execute(self, record: JobRecord) -> None:
        cfg = self._cfg
        check_platform_supported()
        output_dir = recheck_output(cfg, record)
        private_dir = self._job_dir / record.job_id
        private_dir.mkdir(parents=True, exist_ok=True)
        try:
            output_dir.mkdir(parents=True, exist_ok=True)
        except OSError as exc:
            raise McpError("OUTPUT_DENIED", f"출력 폴더를 만들 수 없습니다: {exc}") from exc
        base, _ = resolve_cli_executable(cfg)
        busy_seen = False
        com_unavailable_seen = False
        for index, input_file in enumerate(record.inputs):
            report_path = private_dir / f"step-{index:03d}.json"
            argv = build_cli_argv(
                base,
                input_path=input_file,
                format_type=record.format,
                output_dir=str(output_dir),
                report_path=str(report_path),
                retry_count=1,
            )
            run = run_cli(
                cfg,
                argv=argv,
                job_private_dir=private_dir / f"step-{index:03d}",
                report_path=report_path,
                timeout_seconds=cfg.job_timeout_seconds,
                step_name="convert",
            )
            if run.timed_out:
                record.timed_out = True
                record.items.append(
                    {
                        "input": input_file,
                        "outputs": [],
                        "status": "실패",
                        "detail": f"제한 시간 초과({cfg.job_timeout_seconds}s)",
                        "retry_count": 0,
                    }
                )
                record.failed_count += 1
                self._persist(record)
                continue
            if is_com_unavailable_failure(run) and run.returncode != 0:
                com_unavailable_seen = True
            if is_busy_failure(run) and run.returncode != 0:
                busy_seen = True
                record.items.append(
                    {
                        "input": input_file,
                        "outputs": [],
                        "status": "실패",
                        "detail": "외부 HwpMate 실행 중 (BUSY_EXTERNAL_INSTANCE)",
                        "retry_count": 0,
                    }
                )
                record.failed_count += 1
                self._persist(record)
                break
            try:
                report = read_cli_report(report_path)
            except McpError:
                record.items.append(
                    {
                        "input": input_file,
                        "outputs": [],
                        "status": "실패",
                        "detail": f"REPORT_MISSING (CLI 종료 코드 {run.returncode})",
                        "retry_count": 0,
                    }
                )
                record.failed_count += 1
                self._persist(record)
                continue
            for item in report["items"]:
                record.items.append(item)
            summary = report["summary"]
            record.success_count += int(summary.get("success_count", 0) or 0)
            record.failed_count += int(summary.get("failed_count", 0) or 0)
            record.skipped_count += int(summary.get("skipped_count", 0) or 0)
            record.canceled_count += int(summary.get("canceled_count", 0) or 0)
            record.warnings.extend(report.get("warnings", []))
            self._persist(record)
        record.finished_at = _utcnow_iso()
        if busy_seen and record.success_count == 0 and record.failed_count > 0:
            record.state = "rejected_busy"
            record.error = {
                "code": "BUSY_EXTERNAL_INSTANCE",
                "message": "HwpMate(GUI 또는 다른 CLI)가 실행 중이어서 작업을 시작하지 못했습니다.",
            }
        elif record.timed_out and record.success_count == 0:
            record.state = "timeout"
            record.error = {"code": "TIMEOUT", "message": "제한 시간 안에 변환이 끝나지 않았습니다."}
        else:
            overall = {
                "summary": {
                    "success_count": record.success_count,
                    "failed_count": record.failed_count,
                    "skipped_count": record.skipped_count,
                    "canceled_count": record.canceled_count,
                }
            }
            state, _ = classify_job(0, overall)
            record.state = state
            if state in ("failed", "partially_failed"):
                if com_unavailable_seen and record.success_count == 0:
                    record.error = {
                        "code": "HANCOM_UNAVAILABLE",
                        "message": "한글 COM 초기화에 실패했습니다. 한컴오피스 한글 설치·실행 상태를 확인하세요.",
                    }
                else:
                    record.error = {"code": "OUTPUT_VALIDATION_FAILED", "message": "일부 파일 변환에 실패했습니다."}
        self._persist(record)


def wait_for_state(
    manager: JobManager, job_id: str, *, timeout_seconds: float = 120.0
) -> JobRecord:
    deadline = time.time() + max(1.0, timeout_seconds)
    while time.time() < deadline:
        record = manager.get(job_id)
        if record.state in TERMINAL_STATES:
            return record
        time.sleep(0.2)
    raise TimeoutError(f"job 대기 시간 초과: {job_id}")
