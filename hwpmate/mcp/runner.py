"""승인된 HwpMate CLI subprocess 실행 (구현 설계서 §5.1).

- 고정 인자 배열만 생성하며 shell=False, 사용자 command 필드는 없다.
- stdout/stderr는 job 로그 파일로 수집하고 MCP stdout으로 전달하지 않는다.
- .py 경로는 허용된 인터프리터(sys.executable) + 공식 진입점으로 실행한다.
"""

from __future__ import annotations

import os
import subprocess
import sys
from dataclasses import dataclass
from pathlib import Path

from .config import McpConfig
from .errors import McpError
from .logging_utils import setup_stderr_logging

logger = setup_stderr_logging()

LOCK_BUSY_MARKER = "이미 실행 중"


@dataclass
class CliRun:
    argv: list[str]
    returncode: int
    timed_out: bool
    stdout_path: str
    stderr_path: str
    report_path: str


def resolve_cli_executable(cfg: McpConfig) -> tuple[list[str], str]:
    raw = str(cfg.cli_executable or "").strip()
    if not raw:
        raise McpError("CLI_NOT_TRUSTED", "CLI 실행 파일이 설정되지 않았습니다 (paths.cli_executable).")
    candidate = Path(raw).expanduser()
    try:
        resolved = candidate.resolve()
    except OSError as exc:
        raise McpError("CLI_NOT_TRUSTED", f"CLI 실행 파일 해석 실패: {exc}") from exc
    if not resolved.is_file():
        raise McpError("CLI_NOT_TRUSTED", f"CLI 실행 파일이 존재하지 않습니다: {resolved}")
    if resolved.suffix.lower() == ".py":
        return [sys.executable, str(resolved)], str(resolved)
    return [str(resolved)], str(resolved)


def build_cli_argv(
    base: list[str],
    *,
    input_path: str,
    format_type: str,
    output_dir: str,
    report_path: str,
    retry_count: int = 1,
) -> list[str]:
    return [
        *base,
        "--input",
        str(input_path),
        "--format",
        str(format_type),
        "--output",
        str(output_dir),
        "--report",
        str(report_path),
        "--retry",
        str(max(0, min(3, int(retry_count)))),
        "--pdf-export-mode",
        "saveas_first",
    ]


def _write_limited(path: Path, data: bytes, limit_kb: int) -> None:
    limit = max(16, int(limit_kb)) * 1024
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, "wb") as handle:
        handle.write(data[:limit])
        if len(data) > limit:
            handle.write(b"\n...[truncated]...")


def run_cli(
    cfg: McpConfig,
    *,
    argv: list[str],
    job_private_dir: Path,
    report_path: Path,
    timeout_seconds: int,
    step_name: str = "convert",
) -> CliRun:
    job_private_dir.mkdir(parents=True, exist_ok=True)
    stdout_path = job_private_dir / f"{step_name}.stdout.log"
    stderr_path = job_private_dir / f"{step_name}.stderr.log"
    try:
        completed = subprocess.run(
            argv,
            shell=False,
            cwd=str(job_private_dir),
            env=sanitized_env(),
            capture_output=True,
            timeout=max(1, int(timeout_seconds)),
        )
        timed_out = False
        returncode = int(completed.returncode)
        stdout_data = bytes(completed.stdout or b"")
        stderr_data = bytes(completed.stderr or b"")
    except subprocess.TimeoutExpired as exc:
        timed_out = True
        returncode = 124
        stdout_data = bytes(exc.stdout or b"") if isinstance(exc.stdout, bytes) else b""
        stderr_data = bytes(exc.stderr or b"") if isinstance(exc.stderr, bytes) else b""
    except OSError as exc:
        raise McpError("INTERNAL", f"CLI 실행 실패: {exc}") from exc
    _write_limited(stdout_path, stdout_data, cfg.max_log_kb)
    _write_limited(stderr_path, stderr_data, cfg.max_log_kb)
    return CliRun(
        argv=list(argv),
        returncode=returncode,
        timed_out=timed_out,
        stdout_path=str(stdout_path),
        stderr_path=str(stderr_path),
        report_path=str(report_path),
    )


def _decode_best_effort(data: bytes) -> str:
    import locale
    import sys

    candidates = ["utf-8", sys.getfilesystemencoding(), locale.getpreferredencoding(False), "cp949", "euc-kr"]
    seen: set[str] = set()
    for encoding in candidates:
        if not encoding or encoding.lower() in seen:
            continue
        seen.add(encoding.lower())
        try:
            return data.decode(encoding)
        except (ValueError, LookupError):
            continue
    return data.decode("utf-8", errors="replace")


def read_stderr_text(run: CliRun) -> str:
    try:
        raw = Path(run.stderr_path).read_bytes()
    except OSError:
        return ""
    return _decode_best_effort(raw)


def is_busy_failure(run: CliRun) -> bool:
    return LOCK_BUSY_MARKER in read_stderr_text(run)


def is_com_unavailable_failure(run: CliRun) -> bool:
    return "COM 초기화 실패" in read_stderr_text(run)


def sanitized_env() -> dict[str, str]:
    """자식 프로세스용 환경: 현재 환경을 그대로 전달 (COM 동작 보장).

    임의 명령 문자열을 만들지 않으므로 argv 고정이 보안 경계다.
    """
    return dict(os.environ)
