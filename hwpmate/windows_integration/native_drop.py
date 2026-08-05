"""네이티브 WM_DROPFILES 필터 (관리자 권한 DnD)."""

from __future__ import annotations

import ctypes
import logging
from ctypes import wintypes
from pathlib import Path
from typing import Any, Callable, ClassVar, List, Optional, Set, Tuple

from PyQt6.QtCore import QAbstractNativeEventFilter

from ..constants import SUPPORTED_EXTENSIONS
from ..logging_config import get_logger

logger = get_logger(__name__)

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
                    logger.debug("WM_DROPFILES 메시지 감지!")
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
