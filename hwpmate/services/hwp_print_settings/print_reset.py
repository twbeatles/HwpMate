"""문서 세션 인쇄 기본값(PrintMethod=0) best-effort 리셋.

Execute(실제 인쇄/PDF 생성)는 하지 않는다.
"""

from __future__ import annotations

from typing import Any

from ...logging_config import get_logger
from .constants import PRINT_METHOD_NORMAL

logger = get_logger(__name__)


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
