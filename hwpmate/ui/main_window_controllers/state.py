from __future__ import annotations

from dataclasses import dataclass, field
from typing import TYPE_CHECKING

from ...models import ConversionSummary, ConversionTask, PlannedConversion
from ...services.hwp_security_session import HwpSecuritySession
from ...workers.conversion_worker import ConversionWorker
from ...workers.file_scan_worker import FileScanWorker

if TYPE_CHECKING:
    from PyQt6.QtCore import QTimer


@dataclass
class MainWindowState:
    """Mutable runtime state owned by MainWindow controllers."""

    tasks: list[ConversionTask] = field(default_factory=list)
    plan: PlannedConversion | None = None
    last_summary: ConversionSummary | None = None
    worker: ConversionWorker | None = None
    is_converting: bool = False
    # 스캔 대기·작업 목록 수립 중 (processEvents 재진입 방지)
    is_planning: bool = False
    conversion_start_time: float | None = None
    scan_worker: FileScanWorker | None = None
    scan_mode: str | None = None
    scan_new_file_count: int = 0
    scan_preview_count: int = 0
    scan_started_at: float | None = None
    # 폴더 미리보기 스캔 결과 캐시 (변환 계획 수립 시 재스캔 방지)
    folder_scan_folder: str = ""
    folder_scan_include_sub: bool = True
    folder_scan_files: list[str] = field(default_factory=list)
    folder_scan_accum: list[str] = field(default_factory=list)
    folder_scan_ready: bool = False
    folder_scan_ready_at: float | None = None
    # 캐시 시점 폴더 mtime·파일 수 (신규 파일/폴더 변경 감지)
    folder_scan_dir_mtime: float | None = None
    folder_scan_file_count: int = 0
    force_kill_pending: bool = False
    close_after_worker: bool = False
    # 계획(스캔 대기·작업 수립) 중 종료 요청. processEvents 재진입으로 창 파괴를 막고
    # start_conversion 이 정리된 뒤 close 를 이어간다.
    close_requested: bool = False
    close_after_plan: bool = False
    drag_drop_initialized: bool = False
    selected_format: str = "PDF"
    # 한글 허용/보안 창 전면화 폴링 타이머 (UI 스레드, ConversionController 소유)
    hwp_foreground_timer: QTimer | None = None
    # 변환 세션 보안/전면화 정책
    security_session: HwpSecuritySession = field(default_factory=HwpSecuritySession)
