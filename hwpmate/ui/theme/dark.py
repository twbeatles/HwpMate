# -*- coding: utf-8 -*-
"""다크 테마.

QSS 본문은 `stylesheet.build_stylesheet` 공용 템플릿에서 생성한다. 이 상수는 아이콘 파일 없이
만든 호환용 문자열이며, 실제 적용은 `ThemeManager.apply_theme` 가 아이콘 경로를 포함해 수행한다.
"""

from .palette import DARK_PALETTE
from .stylesheet import build_stylesheet

DARK_THEME = build_stylesheet(DARK_PALETTE, with_icons=False)
