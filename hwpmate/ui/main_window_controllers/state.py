from __future__ import annotations

from dataclasses import dataclass, field

from ...models import ConversionSummary, ConversionTask, PlannedConversion
from ...workers.conversion_worker import ConversionWorker
from ...workers.file_scan_worker import FileScanWorker


@dataclass
class MainWindowState:
    """Mutable runtime state owned by MainWindow controllers."""

    tasks: list[ConversionTask] = field(default_factory=list)
    plan: PlannedConversion | None = None
    last_summary: ConversionSummary | None = None
    worker: ConversionWorker | None = None
    is_converting: bool = False
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
    force_kill_pending: bool = False
    close_after_worker: bool = False
    drag_drop_initialized: bool = False
    selected_format: str = "PDF"
