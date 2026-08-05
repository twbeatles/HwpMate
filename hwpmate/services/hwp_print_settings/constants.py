"""인쇄 설정·PDF 내보내기 관련 상수."""

from __future__ import annotations

from typing import Callable

# SaveAs 시 인쇄/렌더 경로를 탈 수 있는 형식
PRINT_AFFECTED_FORMATS = frozenset({"PDF", "PNG", "JPG", "BMP", "GIF"})

# 한컴 HPrint.PrintMethod: 0=자동 인쇄(1쪽씩), 4=2쪽 모아찍기 등
PRINT_METHOD_NORMAL = 0

# 알려진 가상 PDF 프린터 (탐지 실패 시 후보)
PDF_PRINTER_NAME_CANDIDATES = (
    "Hancom PDF",
    "Microsoft Print to PDF",
)

# PDF 내보내기 모드
PDF_EXPORT_SAVEAS_FIRST = "saveas_first"
PDF_EXPORT_PRINT_TO_PDF_EX_FIRST = "print_to_pdf_ex_first"
PDF_EXPORT_MODES = frozenset({PDF_EXPORT_SAVEAS_FIRST, PDF_EXPORT_PRINT_TO_PDF_EX_FIRST})

# 내보내기 경로 감사 값
EXPORT_METHOD_SAVEAS_2 = "saveas_2"
EXPORT_METHOD_SAVEAS_3 = "saveas_3"
EXPORT_METHOD_PRINT_TO_PDF_EX = "print_to_pdf_ex"
EXPORT_METHOD_RUN_TO_PDF = "run_to_pdf"

# 유효 PDF 최소 크기 (헤더 포함)
MIN_PDF_BYTES = 8
PDF_MAGIC = b"%PDF"

CancelCheck = Callable[[], bool]
