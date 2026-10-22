"""승인된 HwpMate CLI subprocess 실행 (구현 설계서 §5.1).

- 고정 인자 배열만 생성하며 shell=False, 사용자 command 필드는 없다.
- stdout/stderr는 job 로그 파일로 수집하고 MCP stdout으로 전달하지 않는다.
- .py 경로는 허용된 인터프리터(sys.executable) + 공식 진입점으로 실행한다.
- 제한 시간 초과 시 CLI 프로세스 트리(onefile 부트로더 자식 포함)를 종료하고,
  CLI가 기록한 소유 한글 PID만 이미지명 재확인 후 정리한다 (전역 Hwp.exe 종료 금지).
"""

from __future__ import annotations

import json
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
ERROR_ELEVATION_REQUIRED = 740
# requireAdministrator 매니페스트는 PE 리소스(앞부분)에 있으므로 앞쪽만 검사한다.
_MANIFEST_SCAN_BYTES = 8 * 1024 * 1024
_REQUIRE_ADMIN_MARKER = b'level="requireAdministrator"'
_CREATE_NO_WINDOW = getattr(subprocess, "CREATE_NO_WINDOW", 0x08000000)
_KILL_GRACE_SECONDS = 10


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


def executable_requires_admin(path: Path) -> bool:
    """exe 매니페스트가 requireAdministrator 이면 True (best-effort)."""
    if path.suffix.lower() != ".exe":
        return False
    overlap = len(_REQUIRE_ADMIN_MARKER)
    scanned = 0
    tail = b""
    try:
        with open(path, "rb") as handle:
            while scanned < _MANIFEST_SCAN_BYTES:
                chunk = handle.read(1024 * 1024)
                if not chunk:
                    break
                scanned += len(chunk)
                if _REQUIRE_ADMIN_MARKER in tail + chunk:
                    return True
                tail = chunk[-overlap:]
    except OSError:
        return False
    return False


def current_process_is_elevated() -> bool:
    if os.name != "nt":
        return True
    try:
        import ctypes

        return bool(ctypes.windll.shell32.IsUserAnAdmin())
    except Exception:
        return False


def elevation_message(path: str) -> str:
    return (
        f"CLI 실행 파일이 관리자 권한을 요구합니다: {path}. "
        "MCP 클라이언트(호스트)를 관리자 권한으로 실행하거나, "
        "cli_executable 에 소스 진입점(hwptopdf-hwpx_v4.py)을 지정하세요."
    )


def build_cli_argv(
    base: list[str],
    *,
    input_path: str,
    format_type: str,
    output_dir: str,
    report_path: str,
    retry_count: int = 1,
    owned_pids_path: str = "",
) -> list[str]:
    argv = [
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
    if owned_pids_path:
        argv += ["--owned-pids-file", str(owned_pids_path)]
    return argv


def _write_limited(path: Path, data: bytes, limit_kb: int) -> None:
    limit = max(16, int(limit_kb)) * 1024
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, "wb") as handle:
        handle.write(data[:limit])
        if len(data) > limit:
            handle.write(b"\n...[truncated]...")


def _terminate_process_tree(process: subprocess.Popen) -> None:
    """CLI와 그 자식(onefile 부트로더가 띄운 실제 Python 프로세스 등)을 종료한다."""
    if os.name == "nt":
        try:
            subprocess.run(
                ["taskkill", "/PID", str(process.pid), "/T", "/F"],
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
                check=False,
                timeout=30,
                creationflags=_CREATE_NO_WINDOW,
            )
        except (OSError, subprocess.SubprocessError) as exc:
            logger.warning("CLI 프로세스 트리 종료 실패(직접 종료 시도): %s", exc)
    try:
        process.kill()
    except OSError:
        pass


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
        process = subprocess.Popen(
            argv,
            shell=False,
            cwd=str(job_private_dir),
            env=sanitized_env(),
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
        )
    except OSError as exc:
        if getattr(exc, "winerror", None) == ERROR_ELEVATION_REQUIRED:
            raise McpError("ELEVATION_REQUIRED", elevation_message(argv[0])) from exc
        raise McpError("INTERNAL", f"CLI 실행 실패: {exc}") from exc

    timed_out = False
    try:
        stdout_data, stderr_data = process.communicate(timeout=max(1, int(timeout_seconds)))
        returncode = int(process.returncode)
    except subprocess.TimeoutExpired:
        timed_out = True
        returncode = 124
        _terminate_process_tree(process)
        try:
            stdout_data, stderr_data = process.communicate(timeout=_KILL_GRACE_SECONDS)
        except subprocess.TimeoutExpired:
            # 손자 프로세스가 파이프를 쥐고 있으면 더 기다리지 않는다.
            stdout_data, stderr_data = b"", b""
    _write_limited(stdout_path, bytes(stdout_data or b""), cfg.max_log_kb)
    _write_limited(stderr_path, bytes(stderr_data or b""), cfg.max_log_kb)
    return CliRun(
        argv=list(argv),
        returncode=returncode,
        timed_out=timed_out,
        stdout_path=str(stdout_path),
        stderr_path=str(stderr_path),
        report_path=str(report_path),
    )


def read_owned_pids(path: Path) -> set[int]:
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return set()
    pids = payload.get("pids", []) if isinstance(payload, dict) else []
    result: set[int] = set()
    for value in pids if isinstance(pids, list) else []:
        if isinstance(value, int) and not isinstance(value, bool) and value > 0:
            result.add(value)
    return result


def _live_hwp_pids() -> set[int]:
    if os.name != "nt":
        return set()
    from ..services.hwp_converter.process_snapshot import _snapshot_hwp_pids

    return _snapshot_hwp_pids()


def _terminate_pid(pid: int) -> bool:
    try:
        result = subprocess.run(
            ["taskkill", "/PID", str(pid), "/F"],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            check=False,
            timeout=30,
            creationflags=_CREATE_NO_WINDOW,
        )
    except (OSError, subprocess.SubprocessError):
        return False
    return result.returncode == 0


def cleanup_owned_hwp_processes(path: Path) -> list[int]:
    """CLI가 정리하지 못하고 남긴 소유 한글 PID만 종료한다. 종료한 PID 목록을 반환.

    정상 종료한 CLI는 정리 후 빈 목록을 기록하므로 아무것도 하지 않는다.
    PID 재사용에 대비해 현재 한글 이미지명 프로세스인 PID만 대상으로 한다.
    """
    owned = read_owned_pids(path)
    if not owned:
        return []
    live = _live_hwp_pids()
    killed: list[int] = []
    for pid in sorted(owned & live):
        if _terminate_pid(pid):
            killed.append(pid)
            logger.warning("CLI가 남긴 소유 한글 프로세스를 종료했습니다: PID=%s", pid)
    return killed


def _decode_best_effort(data: bytes) -> str:
    import locale

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
