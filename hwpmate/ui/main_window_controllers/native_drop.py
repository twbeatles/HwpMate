from __future__ import annotations

import ctypes
import traceback
from pathlib import Path
from typing import Any

from PyQt6.QtCore import QTimer
from PyQt6.QtWidgets import QApplication, QMessageBox

from ...constants import FEEDBACK_RESET_DELAY
from ...logging_config import get_logger
from ...path_utils import canonicalize_path
from ...windows_integration import NativeDropFilter, get_native_admin_drag_drop_policy
from .state import MainWindowState

logger = get_logger(__name__)


class NativeDropController:
    """Windows native drag-and-drop bridge for elevated GUI runs."""

    def __init__(self, window: Any, state: MainWindowState) -> None:
        self.window = window
        self.state = state

    def initialize_native_drag_drop(self) -> None:
        if self.state.drag_drop_initialized:
            return

        self.state.drag_drop_initialized = True
        try:
            native_dnd_enabled, native_dnd_reason = get_native_admin_drag_drop_policy()
            if not native_dnd_enabled:
                logger.warning(f"네이티브 드래그 앤 드롭 초기화 건너뜀: {native_dnd_reason}")
                return

            drop_filter = NativeDropFilter.get_instance()
            main_hwnd = int(self.window.winId())
            drop_filter.register_window(main_hwnd)

            try:
                user32 = ctypes.windll.user32
                WNDENUMPROC = ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_void_p, ctypes.c_void_p)

                def enum_callback(child_hwnd: int, _: int) -> bool:
                    try:
                        drop_filter.register_window(child_hwnd)
                    except Exception:
                        pass
                    return True

                callback = WNDENUMPROC(enum_callback)
                user32.EnumChildWindows(main_hwnd, callback, 0)
                logger.debug("자식 윈도우 드래그 앤 드롭 등록 완료")
            except Exception as e:
                logger.debug(f"자식 윈도우 열거 실패 (무시): {e}")

            drop_filter.files_dropped_callback = self.window._on_native_files_dropped

            app = QApplication.instance()
            if app:
                app.installNativeEventFilter(drop_filter)
                logger.info("네이티브 이벤트 필터 설치 완료")

            logger.info("네이티브 드래그 앤 드롭 초기화 완료")
        except Exception as e:
            logger.warning(f"네이티브 드래그 앤 드롭 초기화 중 오류: {e}")
            traceback.print_exc()

    def on_native_files_dropped(self, files: list[str]) -> None:
        if not files:
            return

        normalized = [canonicalize_path(path) for path in files if str(path).strip()]
        if not normalized:
            return

        if self.window.folder_radio.isChecked():
            if len(normalized) == 1 and Path(normalized[0]).is_dir():
                folder = normalized[0]
                self.window.folder_entry.setText(folder)
                self.window.config["last_folder"] = folder
                self.window.config["folder_path"] = folder
                self.window._start_folder_preview_scan(folder)
                if hasattr(self.window, "toast"):
                    self.window.toast.show_message("📁 폴더 드롭을 받아 미리보기 스캔을 시작합니다", "✅")
                return

            QMessageBox.warning(
                self.window,
                "경고",
                "폴더 모드에서는 폴더 1개만 드롭할 수 있습니다.\n파일이나 다중 경로 드롭은 지원하지 않습니다.",
            )
            return

        self.window._add_files(normalized)
        if hasattr(self.window, "drop_area") and self.window.drop_area:
            self.window.drop_area.icon_label.setText("✅")
            self.window.drop_area.text_label.setText(f"{len(normalized)}개 경로 스캔 시작")
            QTimer.singleShot(FEEDBACK_RESET_DELAY, self.window.drop_area._reset_appearance)
        if hasattr(self.window, "toast"):
            self.window.toast.show_message(f"📂 {len(normalized)}개 경로를 스캔합니다", "✅")
