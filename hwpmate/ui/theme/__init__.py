"""UI 테마 패키지.

호환: `from hwpmate.ui.theme import ThemeManager`
"""

from __future__ import annotations

from .dark import DARK_THEME
from .light import LIGHT_THEME
from .manager import ThemeManager

__all__ = ["ThemeManager", "DARK_THEME", "LIGHT_THEME"]
