from __future__ import annotations

import ctypes
import logging
import os
import sys
from ctypes import wintypes
from functools import lru_cache
from pathlib import Path
from typing import Any, Callable, ClassVar, List, Optional, Set, Tuple

from PyQt6.QtCore import QAbstractNativeEventFilter

from .constants import SUPPORTED_EXTENSIONS
from .logging_config import get_logger

logger = get_logger(__name__)

NATIVE_DND_DISABLE_ENV = "HWPMATE_DISABLE_NATIVE_DND"
NATIVE_DND_FORCE_ENV = "HWPMATE_FORCE_NATIVE_DND"
WINDOWS_GENERIC_MSG = b"windows_generic_MSG"


def _event_type_bytes(event_type: Any) -> bytes:
    """PyQt6 버전별로 bytes / QByteArray 로 올 수 있는 eventType 을 정규화."""
    if isinstance(event_type, (bytes, bytearray)):
        return bytes(event_type)
    data = getattr(event_type, "data", None)
    if callable(data):
        try:
            raw = data()
            if isinstance(raw, (bytes, bytearray, memoryview)):
                return bytes(raw)
            if isinstance(raw, str):
                return raw.encode("utf-8", errors="replace")
        except Exception:
            pass
    if isinstance(event_type, str):
        return event_type.encode("utf-8", errors="replace")
    try:
        return bytes(event_type)  # type: ignore[arg-type]
    except Exception:
        return str(event_type).encode("utf-8", errors="replace")


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
    if _env_flag(NATIVE_DND_FORCE_ENV):
        return True, f"{NATIVE_DND_FORCE_ENV}=1"

    if _env_flag(NATIVE_DND_DISABLE_ENV):
        return False, f"{NATIVE_DND_DISABLE_ENV}=1"

    if is_running_under_idle():
        return False, "IDLE 환경에서는 셸 재시작/종료 오인 방지를 위해 비활성화"

    return True, "default"

def is_admin() -> bool:
    """관리자 권한 확인"""
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except Exception as e:
        logger.warning(f"관리자 권한 확인 실패: {e}")
        return False


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
            if security_dialogs_only and not is_likely_hwp_security_dialog(int(hwnd)):
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
    from .services.hwp_converter import _snapshot_hwp_pids

    return set(_snapshot_hwp_pids())


def bring_hwp_windows_to_foreground(pids: Optional[Set[int]] = None) -> int:
    """
    한글 관련 창(허용/보안 팝업 포함)을 전면으로 올린다.

    Args:
        pids: 대상 프로세스 PID. None 이면 실행 중인 Hwp/HwpCtrl 전체를 best-effort 로 탐색.

    Returns:
        전면화 시도에 성공한 창 개수 (best-effort).
    """
    try:
        user32 = ctypes.windll.user32
        try:
            # ASFW_ANY = -1 — 다른 프로세스 창 전면화를 허용 (관리자 환경에서 완화)
            user32.AllowSetForegroundWindow(-1)
        except Exception:
            pass

        target_pids = _resolve_hwp_target_pids(pids)
        if not target_pids:
            return 0

        # 메인 편집 창 전체 TOPMOST 반복은 스팸이 되므로 보안/대화상자 위주
        raised = 0
        for hwnd in _list_top_level_hwnds_for_pids(target_pids, security_dialogs_only=True):
            if _raise_hwnd_to_foreground(hwnd):
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
    try:
        user32 = ctypes.windll.user32
        SW_HIDE = 0
        target_pids = _resolve_hwp_target_pids(pids)
        if not target_pids:
            return 0

        hidden = 0
        for hwnd in _list_top_level_hwnds_for_pids(target_pids, security_dialogs_only=False):
            if is_likely_hwp_security_dialog(hwnd):
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
    hidden = hide_hwp_main_windows(pids)
    raised = bring_hwp_windows_to_foreground(pids)
    return hidden, raised


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


def try_accept_hwp_security_dialog(pids: Optional[Set[int]] = None) -> bool:
    """
    한글 보안 승인 대화상자의 「모두 허용」 버튼을 best-effort 로 클릭한다.

    주 경로는 RegisterModule 보안 모듈이다. 모듈 실패 시 남는 UI 만 보조한다.
    버튼 문구/창 구조가 버전마다 달라 실패할 수 있으며, 실패해도 변환을 중단하지 않는다.
    """
    try:
        user32 = ctypes.windll.user32
        target_pids = _resolve_hwp_target_pids(pids)
        if not target_pids:
            return False

        # 메인 편집 창 전체 전면화는 포커스 탈취가 크므로 보안/대화상자 후보만 사용
        parent_hwnds = _list_top_level_hwnds_for_pids(
            target_pids,
            security_dialogs_only=True,
        )
        if not parent_hwnds:
            return False

        # 전면화 후 버튼 탐색
        for parent in parent_hwnds:
            _raise_hwnd_to_foreground(parent)

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
                # 정확 일치 우선, 느슨 매칭은 "모두"+"허용"이 모두 있을 때만
                matched = text in _SECURITY_ACCEPT_BUTTON_TEXTS or (
                    "모두" in text and "허용" in text and len(text) <= 24
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

class NativeDropFilter(QAbstractNativeEventFilter):
    """
    Windows 네이티브 WM_DROPFILES 메시지 처리 필터
    
    관리자 권한으로 실행된 프로세스에서도 드래그 앤 드롭이 작동하도록
    Qt의 OLE 드래그 앤 드롭 대신 Windows Shell의 WM_DROPFILES를 사용합니다.
    """
    
    # 시그널을 위한 싱글톤 객체
    _instance: ClassVar[Optional["NativeDropFilter"]] = None
    files_dropped_callback: Optional[Callable[[List[str]], None]] = None
    
    WM_DROPFILES = 0x0233
    
    def __init__(self) -> None:
        super().__init__()
        self._shell32 = ctypes.windll.shell32
        self.files_dropped_callback = None
        self._registered_hwnds: Set[int] = set()
        self._argtypes_configured = False
        
        # ctypes argtypes를 한 번만 설정
        self._configure_argtypes()
    
    def _configure_argtypes(self) -> None:
        """ctypes 함수 시그니처 설정 (한 번만 실행)"""
        if self._argtypes_configured:
            return
        try:
            self._shell32.DragQueryFileW.argtypes = [ctypes.c_void_p, ctypes.c_uint, ctypes.c_wchar_p, ctypes.c_uint]
            self._shell32.DragQueryFileW.restype = ctypes.c_uint
            self._shell32.DragFinish.argtypes = [ctypes.c_void_p]
            self._shell32.DragFinish.restype = None
            self._argtypes_configured = True
        except Exception as e:
            logger.debug(f"ctypes argtypes 설정 실패: {e}")
        
    @classmethod
    def get_instance(cls) -> "NativeDropFilter":
        if cls._instance is None:
            cls._instance = cls()
        return cls._instance
    
    def register_window(self, hwnd: int) -> bool:
        """윈도우에 드래그 앤 드롭 등록"""
        if hwnd in self._registered_hwnds:
            return True
            
        try:
            shell32 = ctypes.windll.shell32
            user32 = ctypes.windll.user32
            ole32 = ctypes.windll.ole32
            
            # OLE 드래그 앤 드롭 해제 (Qt가 등록했을 수 있음)
            # 이렇게 해야 탐색기가 WM_DROPFILES로 전환함
            try:
                ole32.RevokeDragDrop(hwnd)
                logger.debug(f"OLE 드래그 앤 드롭 해제: HWND={hwnd}")
            except Exception as e:
                logger.debug(f"RevokeDragDrop 실패 (무시 가능): {e}")
            
            # 메시지 필터 허용 (UIPI 우회)
            MSGFLT_ALLOW = 1
            user32.ChangeWindowMessageFilter(self.WM_DROPFILES, MSGFLT_ALLOW)
            user32.ChangeWindowMessageFilter(0x004A, MSGFLT_ALLOW)  # WM_COPYDATA
            user32.ChangeWindowMessageFilter(0x0049, MSGFLT_ALLOW)  # WM_COPYGLOBALDATA
            
            # 윈도우별 필터도 설정
            try:
                user32.ChangeWindowMessageFilterEx(hwnd, self.WM_DROPFILES, MSGFLT_ALLOW, None)
                user32.ChangeWindowMessageFilterEx(hwnd, 0x004A, MSGFLT_ALLOW, None)
                user32.ChangeWindowMessageFilterEx(hwnd, 0x0049, MSGFLT_ALLOW, None)
            except Exception as e:
                logger.debug(f"ChangeWindowMessageFilterEx 실패 (무시): {e}")
            
            # DragAcceptFiles로 WM_DROPFILES 드롭 허용
            shell32.DragAcceptFiles(hwnd, True)
            
            self._registered_hwnds.add(hwnd)
            logger.info(f"네이티브 드래그 앤 드롭 등록 완료: HWND={hwnd}")
            return True
            
        except Exception as e:
            logger.error(f"네이티브 드래그 앤 드롭 등록 실패: {e}")
            return False
    
    def nativeEventFilter(self, eventType: Any, message: Any) -> Tuple[bool, Any]:
        """네이티브 Windows 이벤트 필터.

        런타임 sip 은 두 번째 반환값으로 int 를 기대한다(None 이면 onefile AV 가능).
        PyQt stub 은 voidptr 를 기대할 수 있어 반환 타입은 Any 로 둔다.
        """
        try:
            # Windows 메시지만 처리
            if _event_type_bytes(eventType) != WINDOWS_GENERIC_MSG:
                return False, 0
            if message is None:
                return False, 0

            # message 는 sip.voidptr — 정수 주소로 변환 후 표준 MSG 레이아웃 사용
            msg_ptr = int(message)
            if msg_ptr == 0:
                return False, 0
            msg = wintypes.MSG.from_address(msg_ptr)

            if int(msg.message) == self.WM_DROPFILES:
                if logger.isEnabledFor(logging.DEBUG):
                    logger.debug("WM_DROPFILES 메시지 수신!")
                dropped_files = self._get_dropped_files(int(msg.wParam))

                if dropped_files and self.files_dropped_callback:
                    # 폴더 확장은 여기서 하지 않고 MainWindow 비동기 스캐너에서 처리
                    accepted_inputs = []
                    for raw_path in dropped_files:
                        path_obj = Path(raw_path)
                        if path_obj.is_dir() or raw_path.lower().endswith(SUPPORTED_EXTENSIONS):
                            accepted_inputs.append(raw_path)

                    if accepted_inputs:
                        logger.debug(f"네이티브 드롭 입력: {len(accepted_inputs)}개 경로")
                        self.files_dropped_callback(accepted_inputs)

                # 메시지 처리 완료 (result 는 반드시 int)
                return True, 0

        except Exception as e:
            if logger.isEnabledFor(logging.DEBUG):
                logger.debug(f"nativeEventFilter 오류: {e}")

        return False, 0
    
    def _get_dropped_files(self, hDrop: int) -> List[str]:
        """WM_DROPFILES에서 파일 목록 추출"""
        files: List[str] = []
        hDrop_ptr = ctypes.c_void_p(hDrop)
        try:
            # 드롭된 파일 수 확인 (0xFFFFFFFF = -1 = 파일 수 반환)
            file_count = self._shell32.DragQueryFileW(hDrop_ptr, 0xFFFFFFFF, None, 0)
            
            if logger.isEnabledFor(logging.DEBUG):
                logger.debug(f"드롭된 파일 수: {file_count}")
            
            # 각 파일 경로 추출
            for i in range(file_count):
                length = self._shell32.DragQueryFileW(hDrop_ptr, i, None, 0)
                if length > 0:
                    buffer = ctypes.create_unicode_buffer(length + 1)
                    self._shell32.DragQueryFileW(hDrop_ptr, i, buffer, length + 1)
                    files.append(buffer.value)
                    if logger.isEnabledFor(logging.DEBUG):
                        logger.debug(f"드롭된 파일 {i}: {buffer.value}")
        except Exception as e:
            logger.error(f"드롭 파일 추출 실패: {e}")
            import traceback
            traceback.print_exc()
        finally:
            try:
                self._shell32.DragFinish(hDrop_ptr)
            except Exception:
                pass
        
        return files
