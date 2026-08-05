"""관리자 권한·네이티브 DnD 정책·UIPI 메시지 필터."""

from __future__ import annotations

import ctypes
import os
import sys
from functools import lru_cache
from typing import Optional

from ..logging_config import get_logger

logger = get_logger(__name__)

NATIVE_DND_DISABLE_ENV = "HWPMATE_DISABLE_NATIVE_DND"
NATIVE_DND_FORCE_ENV = "HWPMATE_FORCE_NATIVE_DND"


def _env_flag(name: str) -> bool:
    value = os.environ.get(name, "").strip().lower()
    return value in {"1", "true", "yes", "on"}


def is_running_under_idle() -> bool:
    """IDLE 실행 여부 추정."""
    stdin_module = getattr(getattr(sys, "stdin", None), "__class__", type("", (), {})).__module__
    return "idlelib" in sys.modules or stdin_module.startswith("idlelib")


@lru_cache(maxsize=1)
def get_native_admin_drag_drop_policy() -> tuple[bool, str]:
    """
    관리자용 네이티브 드래그 앤 드롭 활성화 여부와 사유를 반환.
    """
    # monkeypatch 호환: 패키지 속성 경유
    from hwpmate import windows_integration as api

    if _env_flag(NATIVE_DND_FORCE_ENV):
        return True, f"{NATIVE_DND_FORCE_ENV}=1"

    if _env_flag(NATIVE_DND_DISABLE_ENV):
        return False, f"{NATIVE_DND_DISABLE_ENV}=1"

    if api.is_running_under_idle():
        return False, "IDLE 환경에서는 셸 재시작/종료 오인 방지를 위해 비활성화"

    return True, "default"


def is_admin() -> bool:
    """관리자 권한 확인"""
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except Exception as e:
        logger.warning(f"관리자 권한 확인 실패: {e}")
        return False


def enable_drag_drop_for_admin(hwnd: Optional[int] = None) -> None:
    """
    관리자 권한으로 실행 시 드래그 앤 드롭 활성화

    Windows의 UIPI(User Interface Privilege Isolation)로 인해
    일반 사용자 프로세스(탐색기)에서 관리자 프로세스로 드래그 앤 드롭이
    기본적으로 차단됩니다. 이 함수는 메시지 필터를 변경하여 이를 허용합니다.

    Args:
        hwnd: 윈도우 핸들. None이면 전역 필터 사용, 지정하면 해당 윈도우에만 적용
    """
    try:
        # WM_DROPFILES 및 관련 메시지 허용
        WM_DROPFILES = 0x0233
        WM_COPYDATA = 0x004A
        WM_COPYGLOBALDATA = 0x0049
        MSGFLT_ALLOW = 1

        user32 = ctypes.windll.user32

        messages = [WM_DROPFILES, WM_COPYDATA, WM_COPYGLOBALDATA]

        if hwnd is not None:
            # 특정 윈도우에 대한 메시지 필터 (ChangeWindowMessageFilterEx - Windows 7+)
            # 더 정확하고 안정적인 방법
            try:
                for msg in messages:
                    result = user32.ChangeWindowMessageFilterEx(hwnd, msg, MSGFLT_ALLOW, None)
                    if not result:
                        logger.debug(f"ChangeWindowMessageFilterEx 실패: msg={hex(msg)}")
                logger.info(f"윈도우 핸들 {hwnd}에 드래그 앤 드롭 메시지 필터 적용 완료")
            except Exception as e:
                logger.debug(f"ChangeWindowMessageFilterEx 실패, 전역 필터로 대체: {e}")
                # 실패 시 전역 필터로 대체
                for msg in messages:
                    user32.ChangeWindowMessageFilter(msg, MSGFLT_ALLOW)
        else:
            # 전역 메시지 필터 (ChangeWindowMessageFilter)
            try:
                for msg in messages:
                    user32.ChangeWindowMessageFilter(msg, MSGFLT_ALLOW)
                logger.debug("전역 드래그 앤 드롭 메시지 필터 설정 완료")
            except Exception as e:
                logger.debug(f"전역 메시지 필터 설정 실패 (무시 가능): {e}")

    except Exception as e:
        logger.warning(f"드래그 앤 드롭 활성화 실패: {e}")
