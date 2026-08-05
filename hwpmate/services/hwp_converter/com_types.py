"""COM 프로토콜·pywin32 바인딩·실패 결과 정규화."""

from __future__ import annotations

from typing import Any, Optional, Protocol, Tuple

pythoncom: Optional[Any] = None
win32_client: Optional[Any] = None

try:
    import pythoncom as _pythoncom
    from win32com import client as _win32_client

    pythoncom = _pythoncom
    win32_client = _win32_client
    PYWIN32_AVAILABLE = True
except ImportError:
    PYWIN32_AVAILABLE = False


def is_com_failure_result(result: object) -> bool:
    """COM Open/SaveAs 실패 반환값 정규화.

    일부 환경은 False, 일부는 0 을 실패로 돌린다. identity 비교만 쓰면 0 을 놓친다.
    """
    if result is False:
        return True
    if result == 0 and not isinstance(result, bool):
        return True
    return False


class HwpAutomation(Protocol):
    """한글 COM 자동화 객체에서 사용하는 최소 인터페이스."""

    def RegisterModule(self, module_name: str, module_name_alias: str) -> Any: ...
    def SetMessageBoxMode(self, mode: int) -> Any: ...
    def Open(self, path: str, format_name: str, options: str) -> Any: ...
    def SaveAs(self, path: str, format_name: str, options: str = "") -> Any: ...
    def Clear(self, option: int = 0) -> Any: ...
    def Quit(self) -> Any: ...


def require_pywin32() -> Tuple[Any, Any]:
    """pywin32 모듈을 보장하고 반환."""
    if pythoncom is None or win32_client is None:
        raise RuntimeError("pywin32가 필요합니다. `pip install pywin32` 후 다시 실행하세요.")
    return pythoncom, win32_client
