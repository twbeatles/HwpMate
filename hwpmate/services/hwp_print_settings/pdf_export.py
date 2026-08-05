"""PrintToPDFEx / RunToPDF PDF 내보내기 (물리 Print Execute 금지)."""

from __future__ import annotations

from pathlib import Path
from typing import Any, Sequence

from ...logging_config import get_logger
from .constants import (
    EXPORT_METHOD_PRINT_TO_PDF_EX,
    EXPORT_METHOD_RUN_TO_PDF,
    MIN_PDF_BYTES,
    PDF_MAGIC,
    PDF_PRINTER_NAME_CANDIDATES,
    PRINT_METHOD_NORMAL,
    CancelCheck,
)
from .print_reset import _apply_safe_print_items, _set_param
from .printers import resolve_pdf_printer_candidates

logger = get_logger(__name__)


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
