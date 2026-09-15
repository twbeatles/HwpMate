from __future__ import annotations

import ctypes
import hashlib
import json
import os
import shutil
import subprocess
import sys
import tempfile
import time
from collections.abc import Callable, Iterable
from pathlib import Path
from typing import Any
from uuid import uuid4

from ..constants import (
    UPDATE_ARTIFACT_MAX_BYTES,
    UPDATE_BACKUP_KEEP_COUNT,
    UPDATE_REQUEST_TIMEOUT_SECONDS,
)
from .update_manifest import ReleaseManifest


UPDATE_STATUS_APPLIED = "applied"
UPDATE_STATUS_ROLLED_BACK = "rolled_back"
UPDATE_STATUS_FAILED = "failed"

# 부모 앱(onefile 부트로더 포함) 종료 후 exe 잠금 해제까지의 여유
UPDATE_PARENT_WAIT_SECONDS = 30.0
UPDATE_FILE_UNLOCK_WAIT_SECONDS = 30.0
UPDATE_REPLACE_ATTEMPTS = 20
UPDATE_REPLACE_DELAY_SECONDS = 0.5

_SYNCHRONIZE = 0x00100000
_PROCESS_QUERY_LIMITED_INFORMATION = 0x1000
_WAIT_OBJECT_0 = 0x00000000
_WAIT_TIMEOUT = 0x00000102
_ERROR_INVALID_PARAMETER = 87
_STILL_ACTIVE = 259


class UpdateApplyError(RuntimeError):
    """업데이트 적용 실패. status 로 결과(rolled_back / failed)를 명시한다."""

    def __init__(self, message: str, *, status: str = UPDATE_STATUS_FAILED) -> None:
        super().__init__(message)
        self.status = status


def update_status_for_exception(exc: BaseException) -> str:
    status = getattr(exc, "status", None)
    if status in {UPDATE_STATUS_ROLLED_BACK, UPDATE_STATUS_FAILED}:
        return str(status)
    return UPDATE_STATUS_FAILED


def wait_for_process_exit(pid: int, timeout: float = UPDATE_PARENT_WAIT_SECONDS) -> bool:
    """PID 프로세스가 종료될 때까지 대기. 종료(또는 존재하지 않음)면 True.

    Windows 에서 os.kill(pid, 0) 은 존재 확인이 아니라 CTRL_C_EVENT 전송이며,
    콘솔 없는 프로세스에서는 대상 생존 여부와 무관하게 즉시 OSError 가 나므로 사용하지 않는다.
    """
    if pid <= 0:
        return True
    if os.name != "nt":
        deadline = time.monotonic() + timeout
        while time.monotonic() < deadline:
            try:
                os.kill(pid, 0)
            except OSError:
                return True
            time.sleep(0.2)
        return False

    # 공용 ctypes.windll 의 argtypes 를 바꾸지 않도록 전용 인스턴스를 사용한다.
    kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
    kernel32.OpenProcess.restype = ctypes.c_void_p
    kernel32.GetExitCodeProcess.argtypes = [ctypes.c_void_p, ctypes.POINTER(ctypes.c_ulong)]
    kernel32.WaitForSingleObject.argtypes = [ctypes.c_void_p, ctypes.c_uint32]
    kernel32.CloseHandle.argtypes = [ctypes.c_void_p]
    handle = kernel32.OpenProcess(_SYNCHRONIZE, False, int(pid))
    if handle:
        try:
            result = kernel32.WaitForSingleObject(handle, int(max(0.0, timeout) * 1000))
            return result == _WAIT_OBJECT_0
        finally:
            kernel32.CloseHandle(handle)

    if ctypes.get_last_error() == _ERROR_INVALID_PARAMETER:
        return True  # 이미 종료된 PID

    # SYNCHRONIZE 권한이 없으면 제한 조회 권한으로 종료 코드를 폴링한다.
    deadline = time.monotonic() + timeout
    exit_code = ctypes.c_ulong()
    while time.monotonic() < deadline:
        query = kernel32.OpenProcess(_PROCESS_QUERY_LIMITED_INFORMATION, False, int(pid))
        if not query:
            return ctypes.get_last_error() == _ERROR_INVALID_PARAMETER
        try:
            if not kernel32.GetExitCodeProcess(query, ctypes.byref(exit_code)):
                return False
            if exit_code.value != _STILL_ACTIVE:
                return True
        finally:
            kernel32.CloseHandle(query)
        time.sleep(0.2)
    return False


def wait_for_file_writable(path: Path, timeout: float = UPDATE_FILE_UNLOCK_WAIT_SECONDS) -> bool:
    """실행 중인 exe 이미지 잠금이 풀려 쓰기 모드로 열 수 있을 때까지 대기."""
    deadline = time.monotonic() + timeout
    while True:
        try:
            with open(path, "r+b"):
                return True
        except FileNotFoundError:
            return True
        except OSError:
            if time.monotonic() >= deadline:
                return False
            time.sleep(0.25)


def _replace_with_retry(
    source: Path,
    destination: Path,
    *,
    attempts: int = UPDATE_REPLACE_ATTEMPTS,
    delay: float = UPDATE_REPLACE_DELAY_SECONDS,
) -> None:
    last_error: OSError | None = None
    for attempt in range(max(1, attempts)):
        try:
            os.replace(source, destination)
            return
        except OSError as exc:
            last_error = exc
            if attempt + 1 < attempts:
                time.sleep(delay)
    assert last_error is not None
    raise last_error


def unique_update_backup_path(target: str | Path, version: str) -> Path:
    """이전 실패로 남은 백업이 있어도 충돌하지 않는 백업 경로 ({exe}.v{ver}[.{n}].bak)."""
    target_path = Path(target).resolve()
    candidate = target_path.parent / f"{target_path.name}.v{version}.bak"
    counter = 1
    while candidate.exists():
        candidate = target_path.parent / f"{target_path.name}.v{version}.{counter}.bak"
        counter += 1
    return candidate


def cleanup_update_staging(staging_root: str | Path, *, keep: Iterable[Path] = ()) -> int:
    """스테이징 폴더의 남은 헬퍼 exe·미적용 업데이트 exe 를 best-effort 로 정리한다.

    실행 중인 헬퍼는 삭제가 실패하므로 조용히 건너뛴다.
    """
    root = Path(staging_root)
    if not root.is_dir():
        return 0
    keep_keys = {os.path.normcase(str(Path(path).resolve())) for path in keep}
    removed = 0
    for pattern in ("update-helper-*.exe", "update-*.exe"):
        for path in root.glob(pattern):
            if os.path.normcase(str(path.resolve())) in keep_keys:
                continue
            try:
                path.unlink()
                removed += 1
            except OSError:
                continue
    return removed


def atomic_write_json(path: Path, payload: Any, *, ensure_ascii: bool = False) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temp_path: Path | None = None
    try:
        with tempfile.NamedTemporaryFile(
            "w",
            encoding="utf-8",
            dir=path.parent,
            prefix=f".{path.name}.",
            suffix=".tmp",
            delete=False,
        ) as f:
            temp_path = Path(f.name)
            json.dump(payload, f, ensure_ascii=ensure_ascii, indent=2)
        temp_path.replace(path)
    except Exception:
        if temp_path is not None:
            try:
                temp_path.unlink(missing_ok=True)
            except OSError:
                pass
        raise


def update_result_path(staging_root: str | Path) -> Path:
    return (Path(staging_root).resolve() / "last-update-result.json").resolve()


def write_update_result(path: str | Path, payload: dict[str, object]) -> None:
    data = dict(payload)
    data["status"] = str(data.get("status", "failed") or "failed")
    atomic_write_json(Path(path).resolve(), data, ensure_ascii=False)


def consume_update_result(path: str | Path) -> dict[str, object] | None:
    result_path = Path(path).resolve()
    try:
        data = json.loads(result_path.read_text(encoding="utf-8"))
    except FileNotFoundError:
        return None
    except Exception:
        try:
            result_path.unlink(missing_ok=True)
        except OSError:
            pass
        return None
    if not isinstance(data, dict) or str(data.get("status", "")) not in {
        "applied",
        "rolled_back",
        "failed",
    }:
        try:
            result_path.unlink(missing_ok=True)
        except OSError:
            pass
        return None
    try:
        result_path.unlink(missing_ok=True)
    except OSError:
        pass
    return data


def resolve_update_staging_root(
    *,
    custom_root: str | Path | None = None,
) -> Path:
    if custom_root is not None:
        return Path(custom_root).resolve()
    local_app_data = os.environ.get("LOCALAPPDATA", "")
    if local_app_data:
        return (Path(local_app_data) / "HwpMate" / "updates").resolve()
    return (Path.home() / ".hwpmate" / "updates").resolve()


def prepare_staged_update(
    manifest: ReleaseManifest,
    *,
    chunks: Iterable[bytes],
    staging_root: str | Path,
    approve: Callable[[ReleaseManifest, Path], bool] | None = None,
    progress_callback: Callable[[int, int], None] | None = None,
    cancel_check: Callable[[], bool] | None = None,
) -> Path | None:
    root = Path(staging_root).resolve()
    root.mkdir(parents=True, exist_ok=True)
    staged = root / f"update-{manifest.version}-{uuid4().hex}.exe"
    digest = hashlib.sha256()
    total = 0
    cancelled = False
    try:
        with open(staged, "xb") as handle:
            for chunk in chunks:
                if cancel_check is not None and cancel_check():
                    cancelled = True
                    break
                if not isinstance(chunk, bytes):
                    raise TypeError("업데이트 청크는 바이트 형식이어야 합니다.")
                total += len(chunk)
                if total > manifest.artifact_size or total > int(
                    UPDATE_ARTIFACT_MAX_BYTES
                ):
                    raise ValueError("업데이트 파일 크기가 매니페스트와 일치하지 않습니다.")
                digest.update(chunk)
                handle.write(chunk)
                if progress_callback is not None:
                    progress_callback(total, manifest.artifact_size)
            handle.flush()
            os.fsync(handle.fileno())

        if cancelled or (cancel_check is not None and cancel_check()):
            staged.unlink(missing_ok=True)
            return None

        if total != manifest.artifact_size:
            staged.unlink(missing_ok=True)
            raise ValueError("업데이트 파일 크기가 매니페스트와 일치하지 않습니다.")
        if digest.hexdigest().lower() != manifest.artifact_sha256.lower():
            staged.unlink(missing_ok=True)
            raise ValueError("업데이트 파일 SHA-256 해시가 일치하지 않습니다.")
        if approve is not None and not approve(manifest, staged):
            staged.unlink(missing_ok=True)
            return None
        return staged
    except Exception:
        staged.unlink(missing_ok=True)
        raise



def stream_update_artifact(
    manifest: ReleaseManifest,
    *,
    cancel_check: Callable[[], bool] | None = None,
) -> Iterable[bytes]:
    from urllib.parse import urlsplit
    from urllib.request import Request, urlopen

    request = Request(
        manifest.artifact_url,
        headers={"User-Agent": "HwpMate-Updater"},
    )
    with urlopen(
        request,
        timeout=float(UPDATE_REQUEST_TIMEOUT_SECONDS),
    ) as response:
        final_url = urlsplit(response.geturl())
        if final_url.scheme.lower() != "https" or not final_url.hostname:
            raise ValueError("업데이트 아티팩트 리다이렉트는 HTTPS를 유지해야 합니다.")
        while True:
            if cancel_check is not None and cancel_check():
                return
            chunk = response.read(1024 * 1024)
            if not chunk:
                return
            yield chunk



def _validate_apply_paths(target: Path, staged: Path, backup: Path) -> None:
    paths = [target.resolve(), staged.resolve(), backup.resolve()]
    if len(set(paths)) != 3:
        raise ValueError("대상, 스테이징, 백업 경로는 서로 달라야 합니다.")
    if target.suffix.lower() != ".exe" or staged.suffix.lower() != ".exe":
        raise ValueError("대상과 스테이징 파일은 반드시 .exe 확장자여야 합니다.")
    if backup.parent != target.parent:
        raise ValueError("백업 파일은 대상 설치 디렉터리에 위치해야 합니다.")
    if not target.is_file() or not staged.is_file():
        raise FileNotFoundError("대상 파일 또는 스테이징 파일이 존재하지 않습니다.")
    if backup.exists():
        raise FileExistsError(f"백업 파일이 이미 존재합니다: {backup}")
    backup.parent.mkdir(parents=True, exist_ok=True)


def cleanup_update_backups(target: str | Path, *, keep_count: int | None = None) -> None:
    target_path = Path(target).resolve()
    keep = max(0, int(UPDATE_BACKUP_KEEP_COUNT if keep_count is None else keep_count))
    candidates: list[tuple[float, Path]] = []
    for backup in target_path.parent.glob(f"{target_path.name}.v*.bak"):
        try:
            candidates.append((backup.stat().st_mtime, backup))
        except OSError:
            continue
    backups = [item for _mtime, item in sorted(candidates, reverse=True)]
    for backup in backups[keep:]:
        try:
            backup.unlink()
        except OSError:
            continue


def apply_staged_update(
    *,
    target: str | Path,
    staged: str | Path,
    backup: str | Path,
    expected_sha256: str | None = None,
    expected_size: int | None = None,
    smoke_runner: Callable[[Path], bool] | None = None,
    replace_attempts: int = UPDATE_REPLACE_ATTEMPTS,
    replace_delay: float = UPDATE_REPLACE_DELAY_SECONDS,
) -> None:
    target_path = Path(target).resolve()
    staged_path = Path(staged).resolve()
    backup_path = Path(backup).resolve()
    _validate_apply_paths(target_path, staged_path, backup_path)
    if expected_size is not None and staged_path.stat().st_size != int(expected_size):
        raise ValueError("교체 전 업데이트 파일 크기가 일치하지 않습니다.")
    if expected_sha256 is not None:
        digest = hashlib.sha256()
        with open(staged_path, "rb") as handle:
            for chunk in iter(lambda: handle.read(1024 * 1024), b""):
                digest.update(chunk)
        if digest.hexdigest().lower() != str(expected_sha256).strip().lower():
            raise ValueError("교체 전 업데이트 파일 해시가 일치하지 않습니다.")
    shutil.copy2(target_path, backup_path)

    try:
        _replace_with_retry(
            staged_path, target_path, attempts=replace_attempts, delay=replace_delay
        )
    except OSError as exc:
        # 교체 전이므로 대상은 이전 버전 그대로다. 백업 사본은 불필요하므로 정리한다
        # (남기면 같은 버전의 다음 업데이트가 FileExistsError 로 계속 실패).
        try:
            backup_path.unlink(missing_ok=True)
        except OSError:
            pass
        raise UpdateApplyError(
            f"실행 파일 교체에 실패해 이전 버전을 유지합니다: {exc}",
            status=UPDATE_STATUS_FAILED,
        ) from exc

    try:
        if smoke_runner is None:
            completed = subprocess.run(
                [str(target_path), "--smoke"],
                timeout=60,
                check=False,
                capture_output=True,
            )
            smoke_ok = completed.returncode == 0
        else:
            smoke_ok = bool(smoke_runner(target_path))
        if not smoke_ok:
            raise RuntimeError("업데이트된 실행 파일의 스모크 검증에 실패했습니다.")
    except Exception as exc:
        try:
            _replace_with_retry(
                backup_path, target_path, attempts=replace_attempts, delay=replace_delay
            )
        except Exception as rollback_exc:
            raise UpdateApplyError(
                "업데이트 적용과 롤백 복구가 모두 실패했습니다. "
                f"백업 파일로 수동 복구가 필요합니다: {backup_path} ({rollback_exc})",
                status=UPDATE_STATUS_FAILED,
            ) from exc
        raise UpdateApplyError(
            f"업데이트 적용 실패로 이전 버전으로 롤백됨: {exc}",
            status=UPDATE_STATUS_ROLLED_BACK,
        ) from exc

    try:
        cleanup_update_backups(target_path)
    except Exception:
        pass


def launch_update_helper(
    *,
    target: str | Path,
    staged: str | Path,
    backup: str | Path,
    parent_pid: int,
    expected_sha256: str,
    expected_size: int,
    result_file: str | Path,
) -> subprocess.Popen[bytes]:
    staged_path = Path(staged).resolve()
    helper_path = staged_path.parent / f"update-helper-{uuid4().hex}.exe"
    shutil.copy2(Path(sys.executable).resolve(), helper_path)
    return subprocess.Popen(
        [
            str(helper_path),
            "--apply-update",
            "--update-target",
            str(Path(target).resolve()),
            "--update-staged",
            str(staged_path),
            "--update-backup",
            str(Path(backup).resolve()),
            "--update-parent-pid",
            str(int(parent_pid)),
            "--update-expected-sha256",
            str(expected_sha256),
            "--update-expected-size",
            str(int(expected_size)),
            "--update-result-file",
            str(Path(result_file).resolve()),
        ],
        close_fds=True,
        creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
    )
