"""한글 COM 인쇄 설정 best-effort 제어.

PDF/이미지 변환은 문서에 남은 인쇄방식(모아찍기 등)을 따를 수 있다.
이 모듈은 변환 직전에 PrintMethod=0(자동/1쪽씩) 등 안전한 기본값을 적용하고,
PDF 전용으로 PrintToPDFEx / RunToPDF 경로를 시도한다.

금지:
- CreateAction("Print").Execute — 물리 프린터로 실제 출력될 수 있음 (감사 Critical)

원본 디스크 파일은 저장하지 않는다 (세션 내 설정만 변경).
"""

from __future__ import annotations

from pathlib import Path
from typing import Any, Callable, Sequence

from ..logging_config import get_logger

logger = get_logger(__name__)

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


def uses_print_settings_control(format_type: str) -> bool:
    return format_type.upper() in PRINT_AFFECTED_FORMATS


def normalize_pdf_export_mode(value: object, default: str = PDF_EXPORT_SAVEAS_FIRST) -> str:
    if isinstance(value, str) and value.strip().lower() in PDF_EXPORT_MODES:
        return value.strip().lower()
    return default


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


def _set_param(target: Any, key: str, value: Any) -> bool:
    """HParameterSet / ActionSet 에 속성 또는 SetItem 으로 값 설정."""
    try:
        if hasattr(target, "SetItem"):
            target.SetItem(key, value)
            return True
    except Exception:
        pass
    try:
        setattr(target, key, value)
        return True
    except Exception:
        return False


def _apply_safe_print_items(pset: Any) -> int:
    """공통 안전 인쇄 항목. 성공한 항목 수를 반환."""
    pairs: list[tuple[str, Any]] = [
        ("PrintMethod", PRINT_METHOD_NORMAL),
        ("NumCopy", 1),
        ("ReverseOrder", 0),
        ("Pause", 0),
        ("Collate", 1),
        ("PrintImage", 1),
        ("PrintDrawObj", 1),
        ("PrintClickHere", 0),
        ("PrintToFile", 0),
        ("UserOrder", 0),
    ]
    applied = 0
    for key, value in pairs:
        if _set_param(pset, key, value):
            applied += 1
    return applied


def apply_default_print_settings(hwp: Any) -> bool:
    """열린 문서의 세션 인쇄 기본값을 1쪽씩(PrintMethod=0)으로 best-effort 리셋.

    Execute(실제 인쇄/PDF 생성)는 하지 않는다.
    SaveAs 경로 전에 호출해 문서에 남은 모아찍기 등이 덜 반영되게 한다.
    """
    if hwp is None:
        return False

    any_ok = False

    # 1) XHwpPrint 프로퍼티 (문서 단위)
    try:
        docs = getattr(hwp, "XHwpDocuments", None)
        if docs is not None:
            doc = docs.Item(0)
            prn = getattr(doc, "XHwpPrint", None)
            if prn is not None:
                for key, value in (
                    ("PrintMethod", PRINT_METHOD_NORMAL),
                    ("NumCopy", 1),
                    ("ReverseOrder", 0),
                ):
                    try:
                        setattr(prn, key, value)
                        any_ok = True
                    except Exception:
                        pass
    except Exception as e:
        logger.debug(f"XHwpPrint 기본값 설정 실패(무시): {e}")

    # 2) HAction + HParameterSet.HPrint GetDefault (Execute 없음)
    try:
        hparam = getattr(hwp, "HParameterSet", None)
        haction = getattr(hwp, "HAction", None)
        if hparam is not None and haction is not None:
            pset = getattr(hparam, "HPrint", None)
            if pset is not None:
                hset = getattr(pset, "HSet", pset)
                for action_id in ("PrintToPDFEx", "Print"):
                    try:
                        haction.GetDefault(action_id, hset)
                        if _apply_safe_print_items(pset) > 0:
                            any_ok = True
                    except Exception as e:
                        logger.debug(f"HAction.GetDefault({action_id}) 실패(무시): {e}")
    except Exception as e:
        logger.debug(f"HParameterSet 인쇄 기본값 설정 실패(무시): {e}")

    # 3) CreateAction("Print") GetDefault + SetItem 만 (Execute 금지 — 물리 인쇄 위험)
    try:
        create_action = getattr(hwp, "CreateAction", None)
        if callable(create_action):
            act: Any = create_action("Print")
            pset: Any = act.CreateSet()
            act.GetDefault(pset)
            if _apply_safe_print_items(pset) > 0:
                any_ok = True
    except Exception as e:
        logger.debug(f"CreateAction(Print) 기본값 설정 실패(무시): {e}")

    if any_ok:
        logger.debug("인쇄 기본값 리셋 적용(PrintMethod=0, best-effort)")
    else:
        logger.debug("인쇄 기본값 리셋 경로를 적용하지 못함(버전/COM 미지원 가능)")
    return any_ok


def is_valid_pdf_file(path: Path, *, min_bytes: int = MIN_PDF_BYTES) -> bool:
    """PDF 매직 바이트와 최소 크기를 검사한다."""
    try:
        if not path.is_file():
            return False
        size = path.stat().st_size
        if size < min_bytes:
            return False
        with path.open("rb") as f:
            header = f.read(len(PDF_MAGIC))
        return header == PDF_MAGIC
    except OSError:
        return False


def remove_incomplete_output(
    path: Path,
    *,
    before_mtime_ns: int | None,
    before_size: int | None,
) -> None:
    """내보내기 실패 후 깨진/부분 산출물을 정리한다.

    기존에 있던 파일을 덮어쓰지 못한 경우(스냅샷과 동일)에는 삭제하지 않는다.
    """
    try:
        if not path.exists():
            return
        st = path.stat()
        if before_mtime_ns is not None and before_size is not None:
            if st.st_mtime_ns == before_mtime_ns and st.st_size == before_size:
                return
        # 신규 생성이거나 내용이 바뀌었는데 유효 PDF가 아니면 제거
        if before_mtime_ns is None or not is_valid_pdf_file(path):
            path.unlink(missing_ok=True)
            logger.debug(f"불완전 PDF 산출물 정리: {path}")
    except OSError as e:
        logger.debug(f"불완전 PDF 정리 실패(무시): {e}")


def try_export_pdf_via_print_to_pdf_ex(
    hwp: Any,
    output_path: str | Path,
    *,
    cancel_check: CancelCheck | None = None,
    printer_names: Sequence[str] | None = None,
    max_printer_attempts: int = 2,
) -> tuple[bool, str | None]:
    """PrintToPDFEx 또는 RunToPDF 로 PDF 생성 (PrintMethod=0).

    Returns:
        (성공 여부, export_method 또는 None)

    물리 Print Execute 는 사용하지 않는다.
    """
    if hwp is None:
        return False, None

    def _cancelled() -> bool:
        if cancel_check is None:
            return False
        try:
            return bool(cancel_check())
        except Exception:
            return False

    output = Path(output_path)
    output_str = str(output)
    parent = output.parent
    try:
        parent.mkdir(parents=True, exist_ok=True)
    except OSError as e:
        logger.debug(f"PrintToPDFEx 출력 폴더 생성 실패: {e}")
        return False, None

    before_mtime_ns: int | None = None
    before_size: int | None = None
    if output.exists():
        try:
            st = output.stat()
            before_mtime_ns = st.st_mtime_ns
            before_size = st.st_size
        except OSError:
            pass

    def _output_is_success() -> bool:
        if not output.exists():
            return False
        try:
            st = output.stat()
        except OSError:
            return False
        if st.st_size < MIN_PDF_BYTES:
            return False
        if before_mtime_ns is not None:
            if st.st_mtime_ns == before_mtime_ns and st.st_size == (before_size or 0):
                return False
        return is_valid_pdf_file(output)

    candidates = list(printer_names) if printer_names else resolve_pdf_printer_candidates()
    if max_printer_attempts > 0:
        candidates = candidates[:max_printer_attempts]
    if not candidates:
        candidates = list(PDF_PRINTER_NAME_CANDIDATES[:max_printer_attempts or 2])

    # --- HAction PrintToPDFEx only (no CreateAction Print Execute) ---
    try:
        hparam = getattr(hwp, "HParameterSet", None)
        haction = getattr(hwp, "HAction", None)
        if hparam is not None and haction is not None:
            pset = getattr(hparam, "HPrint", None)
            if pset is not None:
                hset = getattr(pset, "HSet", pset)
                for printer_name in candidates:
                    if _cancelled():
                        remove_incomplete_output(
                            output,
                            before_mtime_ns=before_mtime_ns,
                            before_size=before_size,
                        )
                        return False, None
                    try:
                        haction.GetDefault("PrintToPDFEx", hset)
                        _apply_safe_print_items(pset)
                        _set_param(pset, "FileName", output_str)
                        _set_param(pset, "filename", output_str)
                        _set_param(pset, "PrinterName", printer_name)
                        result = haction.Execute("PrintToPDFEx", hset)
                        if _output_is_success():
                            logger.debug(
                                f"PrintToPDFEx 성공: printer={printer_name!r}, "
                                f"result={result!r}, path={output_str}"
                            )
                            return True, EXPORT_METHOD_PRINT_TO_PDF_EX
                        logger.debug(
                            f"PrintToPDFEx 실행 후 유효 PDF 없음: printer={printer_name!r}, "
                            f"result={result!r}"
                        )
                        remove_incomplete_output(
                            output,
                            before_mtime_ns=before_mtime_ns,
                            before_size=before_size,
                        )
                    except Exception as e:
                        logger.debug(f"PrintToPDFEx({printer_name!r}) 실패: {e}")
                        remove_incomplete_output(
                            output,
                            before_mtime_ns=before_mtime_ns,
                            before_size=before_size,
                        )
    except Exception as e:
        logger.debug(f"HAction PrintToPDFEx 경로 실패: {e}")

    if _cancelled():
        return False, None

    # --- XHwpPrint.RunToPDF ---
    try:
        docs = getattr(hwp, "XHwpDocuments", None)
        if docs is not None:
            prn = docs.Item(0).XHwpPrint
            try:
                prn.PrintMethod = PRINT_METHOD_NORMAL
            except Exception:
                pass
            try:
                prn.filename = output_str
            except Exception:
                try:
                    prn.FileName = output_str
                except Exception:
                    pass
            for printer_name in candidates:
                if _cancelled():
                    remove_incomplete_output(
                        output,
                        before_mtime_ns=before_mtime_ns,
                        before_size=before_size,
                    )
                    return False, None
                try:
                    prn.PrinterName = printer_name
                except Exception:
                    pass
                try:
                    prn.RunToPDF()
                    if _output_is_success():
                        logger.debug(f"XHwpPrint.RunToPDF 성공: printer={printer_name!r}")
                        return True, EXPORT_METHOD_RUN_TO_PDF
                    remove_incomplete_output(
                        output,
                        before_mtime_ns=before_mtime_ns,
                        before_size=before_size,
                    )
                except Exception as e:
                    logger.debug(f"RunToPDF({printer_name!r}) 실패: {e}")
                    remove_incomplete_output(
                        output,
                        before_mtime_ns=before_mtime_ns,
                        before_size=before_size,
                    )
    except Exception as e:
        logger.debug(f"XHwpPrint.RunToPDF 경로 실패: {e}")

    remove_incomplete_output(
        output,
        before_mtime_ns=before_mtime_ns,
        before_size=before_size,
    )
    return False, None
