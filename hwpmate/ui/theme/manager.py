# -*- coding: utf-8 -*-
"""테마 관리자."""

from __future__ import annotations

from .dark import DARK_THEME
from .light import LIGHT_THEME


class ThemeManager:
    """테마 관리자"""

    DARK_THEME = DARK_THEME
    LIGHT_THEME = LIGHT_THEME

    @staticmethod
    def get_theme(theme_name: str) -> str:
        if theme_name == "dark":
            return ThemeManager.DARK_THEME
        return ThemeManager.LIGHT_THEME
