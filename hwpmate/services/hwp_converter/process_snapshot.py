"""한글 프로세스 Toolhelp 스냅샷 (콘솔 창 플래시 없음)."""

from __future__ import annotations

import ctypes
from ctypes import wintypes

from ...logging_config import get_logger

logger = get_logger(__name__)

HWP_PROCESS_NAMES = {"hwp.exe", "hwpctrl.exe"}
TH32CS_SNAPPROCESS = 0x00000002

# Toolhelp 스냅샷 연속 실패 추적 (모듈 전역, 프로세스 수명 동안)
_snapshot_failure_count = 0
_snapshot_last_error: str | None = None


class _PROCESSENTRY32W(ctypes.Structure):
    _fields_ = [
        ("dwSize", wintypes.DWORD),
        ("cntUsage", wintypes.DWORD),
        ("th32ProcessID", wintypes.DWORD),
        ("th32DefaultHeapID", ctypes.POINTER(ctypes.c_ulong)),
        ("th32ModuleID", wintypes.DWORD),
        ("cntThreads", wintypes.DWORD),
        ("th32ParentProcessID", wintypes.DWORD),
        ("pcPriClassBase", ctypes.c_long),
        ("dwFlags", wintypes.DWORD),
        ("szExeFile", wintypes.WCHAR * 260),
    ]


def get_snapshot_health() -> tuple[int, str | None]:
    """(연속 실패 횟수, 마지막 오류 메시지) — UI/워커 경고용."""
    return _snapshot_failure_count, _snapshot_last_error


def _snapshot_hwp_pids() -> set[int]:
    """현재 실행 중인 한글 관련 프로세스 PID 집합 반환.

    tasklist 서브프로세스 대신 Toolhelp32 를 사용해 콘솔 창 플래시를 원천 차단한다.
    """
    global _snapshot_failure_count, _snapshot_last_error
    try:
        kernel32 = ctypes.windll.kernel32
        snapshot = kernel32.CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)
        if snapshot in (-1, 0xFFFFFFFF):
            _snapshot_failure_count += 1
            _snapshot_last_error = "CreateToolhelp32Snapshot invalid handle"
            return set()

        pids: set[int] = set()
        try:
            entry = _PROCESSENTRY32W()
            entry.dwSize = ctypes.sizeof(_PROCESSENTRY32W)
            if not kernel32.Process32FirstW(snapshot, ctypes.byref(entry)):
                _snapshot_failure_count += 1
                _snapshot_last_error = "Process32FirstW failed"
                return set()
            while True:
                image_name = entry.szExeFile.strip().lower()
                if image_name in HWP_PROCESS_NAMES:
                    pids.add(int(entry.th32ProcessID))
                if not kernel32.Process32NextW(snapshot, ctypes.byref(entry)):
                    break
            # 성공 시 연속 실패 카운터 리셋
            _snapshot_failure_count = 0
            _snapshot_last_error = None
            return pids
        finally:
            kernel32.CloseHandle(snapshot)
    except Exception as e:
        _snapshot_failure_count += 1
        _snapshot_last_error = str(e)
        logger.debug(f"한글 프로세스 스냅샷 수집 실패: {e}")
        return set()
