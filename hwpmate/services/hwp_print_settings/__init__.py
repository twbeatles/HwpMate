"""한글 COM 인쇄 설정 best-effort 제어.

PDF/이미지 변환은 문서에 남은 인쇄방식(모아찍기 등)을 따를 수 있다.
이 모듈은 변환 직전에 PrintMethod=0(자동/1쪽씩) 등 안전한 기본값을 적용하고,
PDF 전용으로 PrintToPDFEx / RunToPDF 경로를 시도한다.

금지:
- CreateAction("Print").Execute — 물리 프린터로 실제 출력될 수 있음 (감사 Critical)

원본 디스크 파일은 저장하지 않는다 (세션 내 설정만 변경).

호환: 기존 `from hwpmate.services.hwp_print_settings import ...` 경로 유지.
"""

from __future__ import annotations

from .constants import (
    EXPORT_METHOD_PRINT_TO_PDF_EX,
    EXPORT_METHOD_RUN_TO_PDF,
    EXPORT_METHOD_SAVEAS_2,
    EXPORT_METHOD_SAVEAS_3,
    MIN_PDF_BYTES,
    PDF_EXPORT_MODES,
    PDF_EXPORT_PRINT_TO_PDF_EX_FIRST,
    PDF_EXPORT_SAVEAS_FIRST,
    PDF_MAGIC,
    PDF_PRINTER_NAME_CANDIDATES,
    PRINT_AFFECTED_FORMATS,
    PRINT_METHOD_NORMAL,
    CancelCheck,
)
from .modes import normalize_pdf_export_mode, uses_print_settings_control
from .pdf_export import (
    is_valid_pdf_file,
    remove_incomplete_output,
    try_export_pdf_via_print_to_pdf_ex,
)
from .print_reset import (
    _apply_safe_print_items,
    _set_param,
    apply_default_print_settings,
)
from .printers import list_installed_printer_names, resolve_pdf_printer_candidates

__all__ = [
    "CancelCheck",
    "EXPORT_METHOD_PRINT_TO_PDF_EX",
    "EXPORT_METHOD_RUN_TO_PDF",
    "EXPORT_METHOD_SAVEAS_2",
    "EXPORT_METHOD_SAVEAS_3",
    "MIN_PDF_BYTES",
    "PDF_EXPORT_MODES",
    "PDF_EXPORT_PRINT_TO_PDF_EX_FIRST",
    "PDF_EXPORT_SAVEAS_FIRST",
    "PDF_MAGIC",
    "PDF_PRINTER_NAME_CANDIDATES",
    "PRINT_AFFECTED_FORMATS",
    "PRINT_METHOD_NORMAL",
    "_apply_safe_print_items",
    "_set_param",
    "apply_default_print_settings",
    "is_valid_pdf_file",
    "list_installed_printer_names",
    "normalize_pdf_export_mode",
    "remove_incomplete_output",
    "resolve_pdf_printer_candidates",
    "try_export_pdf_via_print_to_pdf_ex",
    "uses_print_settings_control",
]
