"""가상 PDF 프린터 탐지."""

from __future__ import annotations

from typing import Sequence

from ...logging_config import get_logger
from .constants import PDF_PRINTER_NAME_CANDIDATES

logger = get_logger(__name__)


def list_installed_printer_names() -> list[str]:
    """설치된 프린터 이름 목록 (win32print 없으면 빈 목록)."""
    try:
        import win32print  # type: ignore[import-untyped]
    except ImportError:
        return []
    names: list[str] = []
    try:
        flags = getattr(win32print, "PRINTER_ENUM_LOCAL", 2) | getattr(
            win32print, "PRINTER_ENUM_CONNECTIONS", 4
        )
        for entry in win32print.EnumPrinters(flags):
            # (flags, description, name, comment) 형태가 일반적
            if len(entry) >= 3 and entry[2]:
                names.append(str(entry[2]))
    except Exception as e:
        logger.debug(f"프린터 목록 열거 실패: {e}")
    return names


def resolve_pdf_printer_candidates(
    preferred: Sequence[str] | None = None,
) -> list[str]:
    """설치된 가상 PDF 프린터를 우선으로 후보 목록을 만든다."""
    preferred_list = list(preferred) if preferred else list(PDF_PRINTER_NAME_CANDIDATES)
    installed = list_installed_printer_names()
    installed_lower = {name.lower(): name for name in installed}

    ordered: list[str] = []
    seen: set[str] = set()

    def _add(name: str) -> None:
        key = name.lower()
        if key in seen:
            return
        seen.add(key)
        # 설치 목록에 있으면 실제 표기 사용
        ordered.append(installed_lower.get(key, name))

    for name in preferred_list:
        if name.lower() in installed_lower:
            _add(name)
    # 설치 목록에서 PDF/XPS 가상 프린터 추가 탐지
    for name in installed:
        lower = name.lower()
        if any(token in lower for token in ("pdf", "xps", "hancom")):
            _add(name)
    # 미설치여도 후보로 한 번 시도 (Enum 실패 환경)
    if not ordered:
        for name in preferred_list:
            _add(name)
    return ordered
