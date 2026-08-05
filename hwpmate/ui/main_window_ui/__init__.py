"""메인 윈도우 UI 빌더 패키지.

호환: from hwpmate.ui.main_window_ui import build_main_window_ui, MainWindowWidgets, MainWindowCallbacks
"""

from __future__ import annotations

from .builder import build_main_window_ui
from .types import MainWindowCallbacks, MainWindowWidgets

__all__ = ["MainWindowCallbacks", "MainWindowWidgets", "build_main_window_ui"]
