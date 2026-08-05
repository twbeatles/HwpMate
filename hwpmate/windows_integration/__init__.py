"""Windows OS 통합 (관리자 DnD · 한글 창 · 보안 대화상자).

호환: `from hwpmate.windows_integration import NativeDropFilter` 등 기존 경로 유지.
서브모듈 간 호출은 패키지 속성을 경유해 테스트 monkeypatch 가 동작한다.
"""

from __future__ import annotations

import ctypes

from .admin import (
    NATIVE_DND_DISABLE_ENV,
    NATIVE_DND_FORCE_ENV,
    _env_flag,
    enable_drag_drop_for_admin,
    get_native_admin_drag_drop_policy,
    is_admin,
    is_running_under_idle,
)
from .hwp_window_control import (
    _BM_CLICK,
    _SECURITY_ACCEPT_BUTTON_TEXTS,
    bring_hwp_windows_to_foreground,
    hide_hwp_main_windows,
    suppress_hwp_ui_flash,
    try_accept_hwp_security_dialog,
)
from .native_drop import (
    WINDOWS_GENERIC_MSG,
    NativeDropFilter,
    _event_type_bytes,
)
from .window_query import (
    _MAIN_WINDOW_TITLE_HINTS,
    _SECURITY_DIALOG_CLASSES,
    _SECURITY_DIALOG_TITLE_HINTS,
    _list_top_level_hwnds_for_pids,
    _raise_hwnd_to_foreground,
    _resolve_hwp_target_pids,
    _window_class_name,
    _window_title,
    is_likely_hwp_security_dialog,
)

__all__ = [
    "NATIVE_DND_DISABLE_ENV",
    "NATIVE_DND_FORCE_ENV",
    "NativeDropFilter",
    "WINDOWS_GENERIC_MSG",
    "_BM_CLICK",
    "_MAIN_WINDOW_TITLE_HINTS",
    "_SECURITY_ACCEPT_BUTTON_TEXTS",
    "_SECURITY_DIALOG_CLASSES",
    "_SECURITY_DIALOG_TITLE_HINTS",
    "_env_flag",
    "_event_type_bytes",
    "_list_top_level_hwnds_for_pids",
    "_raise_hwnd_to_foreground",
    "_resolve_hwp_target_pids",
    "_window_class_name",
    "_window_title",
    "bring_hwp_windows_to_foreground",
    "ctypes",
    "enable_drag_drop_for_admin",
    "get_native_admin_drag_drop_policy",
    "hide_hwp_main_windows",
    "is_admin",
    "is_likely_hwp_security_dialog",
    "is_running_under_idle",
    "suppress_hwp_ui_flash",
    "try_accept_hwp_security_dialog",
]
