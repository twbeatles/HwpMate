"""한글 COM 변환 엔진 패키지.

호환: `from hwpmate.services.hwp_converter import HWPConverter` 등 기존 경로 유지.
테스트 monkeypatch 는 실제 바인딩 모듈(`converter`, `process_snapshot` 등)을 대상으로 한다.
"""

from __future__ import annotations

from .artifact_snapshot import (
    _FileSnapshot,
    _changed_artifacts,
    _iter_candidate_artifact_files,
    _snapshot_artifacts,
    _snapshot_file,
)
from .com_types import (
    PYWIN32_AVAILABLE,
    HwpAutomation,
    is_com_failure_result,
    pythoncom,
    require_pywin32,
    win32_client,
)
from .converter import (
    SECURITY_MODULE_ALIASES,
    _CREATE_NO_WINDOW,
    HWPConverter,
    subprocess,
)
from .process_snapshot import (
    HWP_PROCESS_NAMES,
    TH32CS_SNAPPROCESS,
    _PROCESSENTRY32W,
    _snapshot_failure_count,
    _snapshot_hwp_pids,
    _snapshot_last_error,
    get_snapshot_health,
)
from .progid import get_registered_hwp_progids

# convert_file 테스트 호환: 패키지 속성으로도 노출 (converter 모듈이 실제 사용)
from ...constants import DOCUMENT_LOAD_DELAY as DOCUMENT_LOAD_DELAY  # noqa: F401
from ..hwp_print_settings import (  # noqa: F401
    apply_default_print_settings,
    try_export_pdf_via_print_to_pdf_ex,
)

__all__ = [
    "DOCUMENT_LOAD_DELAY",
    "HWPConverter",
    "HWP_PROCESS_NAMES",
    "HwpAutomation",
    "PYWIN32_AVAILABLE",
    "SECURITY_MODULE_ALIASES",
    "TH32CS_SNAPPROCESS",
    "_CREATE_NO_WINDOW",
    "_FileSnapshot",
    "_PROCESSENTRY32W",
    "_changed_artifacts",
    "_iter_candidate_artifact_files",
    "_snapshot_artifacts",
    "_snapshot_failure_count",
    "_snapshot_file",
    "_snapshot_hwp_pids",
    "_snapshot_last_error",
    "apply_default_print_settings",
    "get_registered_hwp_progids",
    "get_snapshot_health",
    "is_com_failure_result",
    "pythoncom",
    "require_pywin32",
    "subprocess",
    "try_export_pdf_via_print_to_pdf_ex",
    "win32_client",
]
