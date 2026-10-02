# -*- coding: utf-8 -*-
"""테마 관리자."""

from __future__ import annotations

from typing import Optional

from PyQt6.QtGui import QColor, QPalette
from PyQt6.QtWidgets import QApplication, QWidget

from .dark import DARK_THEME
from .light import LIGHT_THEME
from .palette import ThemePalette, palette_for
from .stylesheet import build_stylesheet


def build_qpalette(p: ThemePalette) -> QPalette:
    """QSS 가 닿지 않는 네이티브 그리기(메시지 박스 아이콘 배경, 팝업 프레임 등)용 팔레트."""
    palette = QPalette()
    roles = {
        QPalette.ColorRole.Window: p.window_bg,
        QPalette.ColorRole.WindowText: p.text,
        QPalette.ColorRole.Base: p.input_bg,
        QPalette.ColorRole.AlternateBase: p.surface_alt,
        QPalette.ColorRole.ToolTipBase: p.toast_bg,
        QPalette.ColorRole.ToolTipText: p.toast_text,
        QPalette.ColorRole.PlaceholderText: p.text_disabled,
        QPalette.ColorRole.Text: p.text,
        QPalette.ColorRole.Button: p.surface,
        QPalette.ColorRole.ButtonText: p.text,
        QPalette.ColorRole.BrightText: p.on_accent,
        QPalette.ColorRole.Highlight: p.accent,
        QPalette.ColorRole.HighlightedText: p.on_accent,
        QPalette.ColorRole.Link: p.accent,
        QPalette.ColorRole.Light: p.surface,
        QPalette.ColorRole.Midlight: p.surface_alt,
        QPalette.ColorRole.Mid: p.border,
        QPalette.ColorRole.Dark: p.border_strong,
        QPalette.ColorRole.Shadow: "#000000",
    }
    for role, color in roles.items():
        palette.setColor(role, QColor(color))
    for role in (
        QPalette.ColorRole.WindowText,
        QPalette.ColorRole.Text,
        QPalette.ColorRole.ButtonText,
    ):
        palette.setColor(QPalette.ColorGroup.Disabled, role, QColor(p.text_disabled))
    return palette


class ThemeManager:
    """테마 관리자"""

    DARK_THEME = DARK_THEME
    LIGHT_THEME = LIGHT_THEME

    @staticmethod
    def normalize(theme_name: object) -> str:
        return "dark" if theme_name == "dark" else "light"

    @staticmethod
    def palette(theme_name: str) -> ThemePalette:
        return palette_for(ThemeManager.normalize(theme_name))

    @staticmethod
    def get_theme(theme_name: str) -> str:
        """아이콘(체크 표시·드롭다운 화살표) 경로를 포함한 QSS."""
        return build_stylesheet(ThemeManager.palette(theme_name))

    @staticmethod
    def apply_theme(theme_name: str, window: Optional[QWidget] = None) -> None:
        """앱 전체(대화상자·메뉴·툴팁·트레이 메뉴 포함)에 테마를 적용한다.

        QApplication 이 없으면 창에만 적용한다.
        """
        css = ThemeManager.get_theme(theme_name)
        app = QApplication.instance()
        if isinstance(app, QApplication):
            app.setPalette(build_qpalette(ThemeManager.palette(theme_name)))
            app.setStyleSheet(css)
            if window is not None and window.styleSheet():
                window.setStyleSheet("")
            return
        if window is not None:
            window.setStyleSheet(css)
