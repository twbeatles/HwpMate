"""PDF 내보내기 모드·인쇄 설정 대상 형식 유틸."""

from __future__ import annotations

from .constants import PDF_EXPORT_MODES, PDF_EXPORT_SAVEAS_FIRST, PRINT_AFFECTED_FORMATS


def uses_print_settings_control(format_type: str) -> bool:
    return format_type.upper() in PRINT_AFFECTED_FORMATS


def normalize_pdf_export_mode(value: object, default: str = PDF_EXPORT_SAVEAS_FIRST) -> str:
    if isinstance(value, str) and value.strip().lower() in PDF_EXPORT_MODES:
        return value.strip().lower()
    return default
