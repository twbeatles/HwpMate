from __future__ import annotations

import ctypes
import logging
from pathlib import Path
from typing import Any, Callable, ClassVar, List, Optional, Set, Tuple

from PyQt6.QtCore import QAbstractNativeEventFilter

from .constants import SUPPORTED_EXTENSIONS
from .logging_config import get_logger

logger = get_logger(__name__)

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
    
    # MSG 구조체를 클래스 레벨로 정의 (반복 생성 방지)
    # ctypes.wintypes를 직접 참조
    class _MSG(ctypes.Structure):
        import ctypes.wintypes as wintypes
        _fields_ = [
            ("hwnd", wintypes.HWND),
            ("message", wintypes.UINT),
            ("wParam", wintypes.WPARAM),
            ("lParam", wintypes.LPARAM),
            ("time", wintypes.DWORD),
            ("pt", wintypes.POINT),
        ]
    
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
        """네이티브 Windows 이벤트 필터"""
        try:
            # Windows 메시지만 처리
            if eventType != b"windows_generic_MSG":
                return False, None
            if message is None:
                return False, None
            
            # 클래스 레벨 MSG 구조체 사용 (매번 재생성 방지)
            # message는 sip.voidptr이므로 정수로 변환 후 MSG로 캐스팅
            msg_ptr = int(message)
            msg = ctypes.cast(msg_ptr, ctypes.POINTER(self._MSG)).contents
            
            if msg.message == self.WM_DROPFILES:
                if logger.isEnabledFor(logging.DEBUG):
                    logger.debug("WM_DROPFILES 메시지 수신!")
                dropped_files = self._get_dropped_files(msg.wParam)
                
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
                
                # 메시지 처리 완료
                return True, None
                
        except Exception as e:
            if logger.isEnabledFor(logging.DEBUG):
                logger.debug(f"nativeEventFilter 오류: {e}")
        
        return False, None
    
    def _get_dropped_files(self, hDrop: int) -> List[str]:
        """WM_DROPFILES에서 파일 목록 추출"""
        files: List[str] = []
        try:
            # 미리 초기화된 shell32 사용 (argtypes도 이미 설정됨)
            # hDrop을 c_void_p로 변환
            hDrop_ptr = ctypes.c_void_p(hDrop)
            
            # 드롭된 파일 수 확인 (0xFFFFFFFF = -1 = 파일 수 반환)
            file_count = self._shell32.DragQueryFileW(hDrop_ptr, 0xFFFFFFFF, None, 0)
            
            if logger.isEnabledFor(logging.DEBUG):
                logger.debug(f"드롭된 파일 수: {file_count}")
            
            # 각 파일 경로 추출
            buffer = ctypes.create_unicode_buffer(260)  # MAX_PATH
            for i in range(file_count):
                length = self._shell32.DragQueryFileW(hDrop_ptr, i, buffer, 260)
                if length > 0:
                    files.append(buffer.value)
                    if logger.isEnabledFor(logging.DEBUG):
                        logger.debug(f"드롭된 파일 {i}: {buffer.value}")
            
            # 드롭 핸들 해제
            self._shell32.DragFinish(hDrop_ptr)
            
        except Exception as e:
            logger.error(f"드롭 파일 추출 실패: {e}")
            import traceback
            traceback.print_exc()
        
        return files
