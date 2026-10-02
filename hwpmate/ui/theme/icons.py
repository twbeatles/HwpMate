# -*- coding: utf-8 -*-
"""QSS 서브컨트롤용 작은 아이콘(체크 표시·화살표)을 런타임에 PNG 로 생성한다.

QSS 의 `image: url(...)` 은 파일 경로가 필요하다. 리소스 번들/SVG 플러그인에 의존하지 않도록
QPainter 로 1x/@2x PNG 를 그려 사용자 캐시 폴더에 둔다. Qt 는 `@2x` 파일을 고해상도
화면에서 자동으로 고른다. 생성 실패 시 빈 문자열을 돌려주며, QSS 는 이미지 없이도 동작한다.
"""

from __future__ import annotations

import os
import tempfile
from pathlib import Path
from typing import Callable

from PyQt6.QtCore import QPointF, Qt
from PyQt6.QtGui import QColor, QImage, QPainter, QPainterPath, QPen

from ...logging_config import get_logger

logger = get_logger(__name__)

_ICON_DIR_NAME = "theme-icons"
_cached_dir: Path | None = None


def _icon_dir() -> Path:
    global _cached_dir
    if _cached_dir is not None:
        return _cached_dir
    base = os.environ.get("LOCALAPPDATA")
    candidates = []
    if base:
        candidates.append(Path(base) / "HwpMate" / _ICON_DIR_NAME)
    candidates.append(Path(tempfile.gettempdir()) / "HwpMate" / _ICON_DIR_NAME)
    for candidate in candidates:
        try:
            candidate.mkdir(parents=True, exist_ok=True)
            _cached_dir = candidate
            return candidate
        except OSError:
            continue
    raise OSError("테마 아이콘 폴더를 만들 수 없습니다")


def _draw_check(painter: QPainter, size: float) -> None:
    path = QPainterPath()
    path.moveTo(QPointF(size * 0.24, size * 0.52))
    path.lineTo(QPointF(size * 0.43, size * 0.70))
    path.lineTo(QPointF(size * 0.77, size * 0.32))
    painter.drawPath(path)


def _draw_chevron_down(painter: QPainter, size: float) -> None:
    path = QPainterPath()
    path.moveTo(QPointF(size * 0.25, size * 0.38))
    path.lineTo(QPointF(size * 0.50, size * 0.63))
    path.lineTo(QPointF(size * 0.75, size * 0.38))
    painter.drawPath(path)


def _draw_chevron_up(painter: QPainter, size: float) -> None:
    path = QPainterPath()
    path.moveTo(QPointF(size * 0.25, size * 0.62))
    path.lineTo(QPointF(size * 0.50, size * 0.37))
    path.lineTo(QPointF(size * 0.75, size * 0.62))
    painter.drawPath(path)


_SHAPES: dict[str, tuple[Callable[[QPainter, float], None], float]] = {
    # name: (draw func, stroke width ratio)
    "check": (_draw_check, 0.13),
    "chevron-down": (_draw_chevron_down, 0.12),
    "chevron-up": (_draw_chevron_up, 0.12),
}


def _render(shape: str, color: str, size: int, path: Path) -> None:
    draw, stroke_ratio = _SHAPES[shape]
    image = QImage(size, size, QImage.Format.Format_ARGB32_Premultiplied)
    image.fill(Qt.GlobalColor.transparent)
    painter = QPainter(image)
    try:
        painter.setRenderHint(QPainter.RenderHint.Antialiasing, True)
        pen = QPen(QColor(color))
        pen.setWidthF(max(1.5, size * stroke_ratio))
        pen.setCapStyle(Qt.PenCapStyle.RoundCap)
        pen.setJoinStyle(Qt.PenJoinStyle.RoundJoin)
        painter.setPen(pen)
        painter.setBrush(Qt.BrushStyle.NoBrush)
        draw(painter, float(size))
    finally:
        painter.end()
    if not image.save(str(path), "PNG"):
        raise OSError(f"아이콘 저장 실패: {path}")


def icon_url(shape: str, color: str, size: int) -> str:
    """QSS `url()` 에 넣을 경로(슬래시 구분)를 반환한다. 실패 시 빈 문자열."""
    try:
        folder = _icon_dir()
        stem = f"{shape}-{color.lstrip('#').lower()}-{size}"
        base_path = folder / f"{stem}.png"
        hidpi_path = folder / f"{stem}@2x.png"
        if not base_path.is_file():
            _render(shape, color, size, base_path)
        if not hidpi_path.is_file():
            _render(shape, color, size * 2, hidpi_path)
        return base_path.as_posix()
    except Exception as exc:  # 아이콘은 장식 요소 — 실패해도 테마 적용은 계속한다.
        logger.debug(f"테마 아이콘 생성 실패 ({shape}, {color}): {exc}")
        return ""
