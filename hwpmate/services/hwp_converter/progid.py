"""한글 COM ProgID 레지스트리 조회."""

from __future__ import annotations

from ...constants import HWP_PROGIDS


def get_registered_hwp_progids() -> list[str]:
    """레지스트리에서 확인 가능한 한글 COM ProgID 목록을 반환."""
    try:
        import winreg
    except ImportError:
        return []

    registered: list[str] = []
    for progid in HWP_PROGIDS:
        try:
            with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, progid):
                registered.append(progid)
        except OSError:
            continue
    return registered
