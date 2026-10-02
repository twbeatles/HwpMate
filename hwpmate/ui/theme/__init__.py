"""UI 테마 패키지.

호환: `from hwpmate.ui.theme import ThemeManager`
"""

from __future__ import annotations

from .dark import DARK_THEME
from .light import LIGHT_THEME
from .manager import ThemeManager
from .palette import DARK_PALETTE, LIGHT_PALETTE, ThemePalette, palette_for

__all__ = [
    "ThemeManager",
    "DARK_THEME",
    "LIGHT_THEME",
    "ThemePalette",
    "DARK_PALETTE",
    "LIGHT_PALETTE",
    "palette_for",
]
