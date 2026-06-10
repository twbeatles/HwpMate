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
    force_kill_pending: bool = False
    close_after_worker: bool = False
    drag_drop_initialized: bool = False
    selected_format: str = "PDF"
