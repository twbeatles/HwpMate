"""HWND 제목/클래스 조회·보안 대화상자 판별·열거."""

from __future__ import annotations

import ctypes
from ctypes import wintypes
from typing import Optional, Set

from ..logging_config import get_logger

logger = get_logger(__name__)

# 보안/허용 대화상자 제목 힌트.
# "한글"/"한/글"/"hwp" 는 메인 편집 창 제목(예: "빈 문서 1 - 한글")에도 들어가
# 전면화·숨김 판별이 깨지므로 넣지 않는다.
_SECURITY_DIALOG_TITLE_HINTS = (
    "보안",
    "허용",
    "접근",
    "automation",
    "allow all",
    "allow",
    "permission",
    "파일 접근",
    "매크로",
    "신뢰",
)

_SECURITY_DIALOG_CLASSES = frozenset({"#32770", "HwpDialog", "ThunderRT6FormDC"})

# 메인 편집 창 제목 패턴 (숨김 대상 판별 보조)
_MAIN_WINDOW_TITLE_HINTS = (
    "빈 문서",
    " - 한글",
    " - 한/글",
    ".hwp",
    ".hwpx",
)


def _window_title(hwnd: int) -> str:
    user32 = ctypes.windll.user32
    length = user32.GetWindowTextLengthW(hwnd)
    if length <= 0:
        return ""
    buf = ctypes.create_unicode_buffer(length + 1)
    user32.GetWindowTextW(hwnd, buf, length + 1)
    return buf.value.strip()


def _window_class_name(hwnd: int) -> str:
    user32 = ctypes.windll.user32
    cls = ctypes.create_unicode_buffer(64)
    user32.GetClassNameW(hwnd, cls, 64)
    return cls.value


def is_likely_hwp_security_dialog(hwnd: int) -> bool:
    """보안·허용·모달 대화상자 후보인지 판별 (메인 편집 창은 False).

    전면화·자동 클릭 대상과, 메인 창 숨김 시 제외 대상을 공통으로 쓴다.
    """
    try:
        user32 = ctypes.windll.user32
        if not user32.IsWindow(hwnd):
            return False

        class_name = _window_class_name(hwnd)
        class_l = class_name.lower()
        if class_name in _SECURITY_DIALOG_CLASSES or "dialog" in class_l:
            return True

        title = _window_title(hwnd)
        # 제목 없는 top-level 은 한컴 모달 후보로 취급 (숨기지 않고 전면화 후보)
        if not title:
            return True

        title_l = title.lower()
        if any(hint in title_l for hint in _SECURITY_DIALOG_TITLE_HINTS):
            return True

        # 메인 편집 창 패턴이면 보안 대화상자가 아님
        if any(hint.lower() in title_l for hint in _MAIN_WINDOW_TITLE_HINTS):
            return False

        return False
    except Exception:
        return False


def _list_top_level_hwnds_for_pids(
    target_pids: Set[int],
    *,
    security_dialogs_only: bool = False,
) -> list[int]:
    """지정 PID 소유의 top-level 윈도우 HWND 목록을 반환."""
    if not target_pids:
        return []

    # monkeypatch 호환
    from hwpmate import windows_integration as api

    user32 = ctypes.windll.user32
    hwnds: list[int] = []

    @ctypes.WINFUNCTYPE(ctypes.c_bool, wintypes.HWND, wintypes.LPARAM)
    def enum_proc(hwnd: int, _lparam: int) -> bool:
        try:
            if not user32.IsWindow(hwnd):
                return True
            # 자식 컨트롤(WS_CHILD)은 제외. 소유된 대화상자(허용/보안 팝업)는 포함.
            GWL_STYLE = -16
            WS_CHILD = 0x40000000
            style = user32.GetWindowLongW(hwnd, GWL_STYLE)
            if style & WS_CHILD:
                return True
            pid = wintypes.DWORD(0)
            user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
            if int(pid.value) not in target_pids:
                return True
            if security_dialogs_only and not api.is_likely_hwp_security_dialog(int(hwnd)):
                return True
            hwnds.append(int(hwnd))
        except Exception:
            pass
        return True

    try:
        user32.EnumWindows(enum_proc, 0)
    except Exception as e:
        logger.debug(f"EnumWindows 실패: {e}")
    return hwnds


def _raise_hwnd_to_foreground(hwnd: int) -> bool:
    """단일 HWND 를 전면으로 올린다. 실패해도 False 만 반환."""
    user32 = ctypes.windll.user32
    SW_RESTORE = 9
    HWND_TOPMOST = -1
    HWND_NOTOPMOST = -2
    SWP_NOMOVE = 0x0002
    SWP_NOSIZE = 0x0001
    SWP_SHOWWINDOW = 0x0040

    try:
        if user32.IsIconic(hwnd):
            user32.ShowWindow(hwnd, SW_RESTORE)
        else:
            user32.ShowWindow(hwnd, 5)  # SW_SHOW

        flags = SWP_NOMOVE | SWP_NOSIZE | SWP_SHOWWINDOW
        user32.SetWindowPos(hwnd, HWND_TOPMOST, 0, 0, 0, 0, flags)
        user32.SetWindowPos(hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, flags)
        user32.BringWindowToTop(hwnd)
        user32.SetForegroundWindow(hwnd)
        return True
    except Exception as e:
        logger.debug(f"HWND 전면화 실패: hwnd={hwnd}, {e}")
        return False


def _resolve_hwp_target_pids(pids: Optional[Set[int]] = None) -> Set[int]:
    if pids is not None:
        return {int(p) for p in pids if int(p) > 0}
    # 순환 import 방지: 프로세스 스냅샷은 hwp_converter 쪽 (Toolhelp, 콘솔 없음)
    from ..services.hwp_converter import _snapshot_hwp_pids

    return set(_snapshot_hwp_pids())
