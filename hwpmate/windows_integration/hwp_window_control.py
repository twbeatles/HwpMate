"""한글 창 전면화·숨김·보안 대화상자 자동 클릭."""

from __future__ import annotations

import ctypes
from ctypes import wintypes
from typing import Optional, Set

from ..logging_config import get_logger

logger = get_logger(__name__)

# 보안 승인 대화상자 버튼 문구 후보 (버전·언어 표기 차이)
_SECURITY_ACCEPT_BUTTON_TEXTS = (
    "모두 허용",
    "모두 허용(&A)",
    "모두허용",
    "Allow all",
    "Allow All",
    "&Allow all",
)

_BM_CLICK = 0x00F5


def bring_hwp_windows_to_foreground(pids: Optional[Set[int]] = None) -> int:
    """
    한글 관련 창(허용/보안 팝업 포함)을 전면으로 올린다.

    Args:
        pids: 대상 프로세스 PID. None 이면 실행 중인 Hwp/HwpCtrl 전체를 best-effort 로 탐색.

    Returns:
        전면화 시도에 성공한 창 개수 (best-effort).
    """
    from hwpmate import windows_integration as api

    try:
        user32 = ctypes.windll.user32
        try:
            # ASFW_ANY = -1 — 다른 프로세스 창 전면화를 허용 (관리자 환경에서 완화)
            user32.AllowSetForegroundWindow(-1)
        except Exception:
            pass

        target_pids = api._resolve_hwp_target_pids(pids)
        if not target_pids:
            return 0

        # 메인 편집 창 전체 TOPMOST 반복은 스팸이 되므로 보안/대화상자 위주
        raised = 0
        for hwnd in api._list_top_level_hwnds_for_pids(target_pids, security_dialogs_only=True):
            if api._raise_hwnd_to_foreground(hwnd):
                raised += 1

        if raised:
            logger.debug(f"한글 보안/대화 창 전면화: pids={sorted(target_pids)}, raised={raised}")
        return raised
    except Exception as e:
        logger.debug(f"한글 창 전면화 전체 실패: {e}")
        return 0


def hide_hwp_main_windows(pids: Optional[Set[int]] = None) -> int:
    """
    앱 소유 한글 프로세스의 메인 편집 창을 SW_HIDE 로 숨긴다.

    보안·허용 대화상자는 숨기지 않는다 (전면화/자동 클릭 경로 유지).
    Dispatch 기동 순간 플래시 자체는 막지 못하지만, 변환 중 메인 UI 노출을 줄인다.
    """
    from hwpmate import windows_integration as api

    try:
        user32 = ctypes.windll.user32
        SW_HIDE = 0
        target_pids = api._resolve_hwp_target_pids(pids)
        if not target_pids:
            return 0

        hidden = 0
        for hwnd in api._list_top_level_hwnds_for_pids(target_pids, security_dialogs_only=False):
            if api.is_likely_hwp_security_dialog(hwnd):
                continue
            try:
                if not user32.IsWindowVisible(hwnd):
                    continue
                # ShowWindow 반환값은 "이전에 보였는지" 이므로 숨김 성공 여부와 무관.
                # 호출 후 비가시이면 집계한다.
                user32.ShowWindow(hwnd, SW_HIDE)
                if not user32.IsWindowVisible(hwnd):
                    hidden += 1
            except Exception as e:
                logger.debug(f"HWND 숨김 실패: hwnd={hwnd}, {e}")
        if hidden:
            logger.debug(f"한글 메인 창 숨김: pids={sorted(target_pids)}, hidden={hidden}")
        return hidden
    except Exception as e:
        logger.debug(f"한글 메인 창 숨김 전체 실패: {e}")
        return 0


def suppress_hwp_ui_flash(pids: Optional[Set[int]] = None) -> tuple[int, int]:
    """
    메인 창 숨김 + 보안 대화상자 전면화를 한 번에 수행.

    Returns:
        (hidden_count, raised_security_count)
    """
    from hwpmate import windows_integration as api

    hidden = api.hide_hwp_main_windows(pids)
    raised = api.bring_hwp_windows_to_foreground(pids)
    return hidden, raised


def try_accept_hwp_security_dialog(pids: Optional[Set[int]] = None) -> bool:
    """
    한글 보안 승인 대화상자의 「모두 허용」 버튼을 best-effort 로 클릭한다.

    주 경로는 RegisterModule 보안 모듈이다. 모듈 실패 시 남는 UI 만 보조한다.
    버튼 문구/창 구조가 버전마다 달라 실패할 수 있으며, 실패해도 변환을 중단하지 않는다.
    """
    from hwpmate import windows_integration as api

    try:
        user32 = ctypes.windll.user32
        target_pids = api._resolve_hwp_target_pids(pids)
        if not target_pids:
            return False

        # 메인 편집 창 전체 전면화는 포커스 탈취가 크므로 보안/대화상자 후보만 사용
        parent_hwnds = api._list_top_level_hwnds_for_pids(
            target_pids,
            security_dialogs_only=True,
        )
        if not parent_hwnds:
            return False

        # 전면화 후 버튼 탐색
        for parent in parent_hwnds:
            api._raise_hwnd_to_foreground(parent)

        clicked = False

        @ctypes.WINFUNCTYPE(ctypes.c_bool, wintypes.HWND, wintypes.LPARAM)
        def enum_child(hwnd: int, _lparam: int) -> bool:
            nonlocal clicked
            if clicked:
                return False
            try:
                length = user32.GetWindowTextLengthW(hwnd)
                if length <= 0 or length > 64:
                    return True
                buf = ctypes.create_unicode_buffer(length + 1)
                user32.GetWindowTextW(hwnd, buf, length + 1)
                text = buf.value.strip()
                # 오클릭 방지를 위해 정확 일치만 허용 (공백 정규화)
                normalized = " ".join(text.split())
                matched = (
                    text in _SECURITY_ACCEPT_BUTTON_TEXTS
                    or normalized in _SECURITY_ACCEPT_BUTTON_TEXTS
                )
                if not matched:
                    return True
                # 버튼 클래스인지 확인 (오탐 축소)
                cls = ctypes.create_unicode_buffer(64)
                user32.GetClassNameW(hwnd, cls, 64)
                class_name = cls.value.lower()
                if "button" in class_name or class_name.startswith("button"):
                    user32.SendMessageW(hwnd, _BM_CLICK, 0, 0)
                    clicked = True
                    logger.info(f"한글 보안 대화상자 '모두 허용' 자동 클릭 시도: '{text}'")
                    return False
            except Exception:
                pass
            return True

        for parent in parent_hwnds:
            user32.EnumChildWindows(parent, enum_child, 0)
            if clicked:
                break

        return clicked
    except Exception as e:
        logger.debug(f"모두 허용 자동 클릭 실패: {e}")
        return False
