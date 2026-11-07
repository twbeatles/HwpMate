"""한글 프로세스 Toolhelp 스냅샷 (콘솔 창 플래시 없음)."""

from __future__ import annotations

import ctypes
from ctypes import wintypes
from typing import Iterable

from ...logging_config import get_logger

logger = get_logger(__name__)

HWP_PROCESS_NAMES = {"hwp.exe", "hwpctrl.exe"}
# COM LocalServer 를 띄우는 DcomLaunch 서비스 호스트
COM_LAUNCHER_PROCESS_NAMES = {"svchost.exe"}
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


def _snapshot_process_table() -> dict[int, tuple[str, int]] | None:
    """전체 프로세스 표 {pid: (소문자 이미지명, 부모 PID)}. 실패 시 None.

    tasklist 서브프로세스 대신 Toolhelp32 를 사용해 콘솔 창 플래시를 원천 차단한다.
    """
    global _snapshot_failure_count, _snapshot_last_error
    try:
        kernel32 = ctypes.windll.kernel32
        snapshot = kernel32.CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)
        if snapshot in (-1, 0xFFFFFFFF):
            _snapshot_failure_count += 1
            _snapshot_last_error = "CreateToolhelp32Snapshot invalid handle"
            return None

        table: dict[int, tuple[str, int]] = {}
        try:
            entry = _PROCESSENTRY32W()
            entry.dwSize = ctypes.sizeof(_PROCESSENTRY32W)
            if not kernel32.Process32FirstW(snapshot, ctypes.byref(entry)):
                _snapshot_failure_count += 1
                _snapshot_last_error = "Process32FirstW failed"
                return None
            while True:
                table[int(entry.th32ProcessID)] = (
                    entry.szExeFile.strip().lower(),
                    int(entry.th32ParentProcessID),
                )
                if not kernel32.Process32NextW(snapshot, ctypes.byref(entry)):
                    break
            # 성공 시 연속 실패 카운터 리셋
            _snapshot_failure_count = 0
            _snapshot_last_error = None
            return table
        finally:
            kernel32.CloseHandle(snapshot)
    except Exception as e:
        _snapshot_failure_count += 1
        _snapshot_last_error = str(e)
        logger.debug(f"프로세스 스냅샷 수집 실패: {e}")
        return None


def _snapshot_hwp_pids() -> set[int]:
    """현재 실행 중인 한글 관련 프로세스 PID 집합 반환."""
    table = _snapshot_process_table()
    if table is None:
        return set()
    return {pid for pid, (image, _ppid) in table.items() if image in HWP_PROCESS_NAMES}


def select_com_launched_pids(
    candidates: Iterable[int],
    table: dict[int, tuple[str, int]] | None = None,
    *,
    allow_unknown: bool | None = None,
) -> tuple[set[int], set[int]]:
    """후보 한글 PID 중 COM(DCOM 런처)이 띄운 것만 고른다. (채택, 제외)

    COM LocalServer(`hwp.exe -Automation`)는 DcomLaunch 서비스(svchost.exe)가 띄운다.
    사용자가 탐색기·바로가기로 연 한글은 부모가 explorer.exe 등이므로 제외한다.
    채택된 PID 의 자식 한글 프로세스는 함께 채택한다.
    부모를 확인할 수 없는 후보(부모가 이미 종료)는 allow_unknown 으로 정한다.
    None(기본)이면 후보가 하나뿐일 때만 채택한다(기존 추적 동작 유지).
    DcomLaunch 는 세션 내내 살아 있으므로 부모 불명은 보통 COM 실행이 아니다.
    프로세스 표를 얻지 못하면 판정할 수 없으므로 후보를 그대로 돌려준다.
    """
    pending = set(candidates)
    if not pending:
        return set(), set()
    if table is None:
        table = _snapshot_process_table()
    if table is None:
        return pending, set()

    total = len(pending)
    accepted: set[int] = set()
    rejected: set[int] = set()
    unknown: set[int] = set()

    # 이미 종료된 후보는 대상에서 뺀다 (PID 재사용 방지)
    for pid in list(pending):
        info = table.get(pid)
        if info is None or info[0] not in HWP_PROCESS_NAMES:
            pending.discard(pid)
            rejected.add(pid)

    changed = True
    while changed:
        changed = False
        for pid in sorted(pending):
            _image, ppid = table[pid]
            parent = table.get(ppid)
            if ppid in accepted or (parent is not None and parent[0] in COM_LAUNCHER_PROCESS_NAMES):
                accepted.add(pid)
            elif ppid in pending:
                continue  # 부모 후보 판정 후 다시 본다
            elif parent is None:
                unknown.add(pid)
            else:
                rejected.add(pid)
            pending.discard(pid)
            changed = True

    # 부모 체인이 후보끼리만 이어져 결론이 안 난 경우 → 판정 불가
    unknown |= pending
    if unknown:
        if allow_unknown is None:
            allow_unknown = total == 1
        if allow_unknown:
            accepted |= unknown
        else:
            rejected |= unknown
    return accepted, rejected
