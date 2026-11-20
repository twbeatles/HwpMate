"""2026-11 기능 감사(PROJECT_AUDIT.md) 수정 회귀 테스트.

ISSUE-001 소유 한글 PID 판정, ISSUE-002 한글 연결 끊김 즉시 재연결,
ISSUE-004 CLI 리포트 저장 실패 처리, 재순환 시 고아 한글 정리.
"""

from __future__ import annotations

from pathlib import Path
from typing import Any

import pytest

from hwpmate.app import _parse_args, _run_cli_conversion
from hwpmate.models import ConversionTask, PlannedConversion
from hwpmate.services.hwp_converter import HWPConverter
from hwpmate.services.hwp_converter.converter import is_com_disconnect_error
from hwpmate.services.hwp_converter.process_snapshot import select_com_launched_pids
from hwpmate.services.task_planner import TaskPlanner
from hwpmate.workers.conversion_worker import ConversionWorker
from hwpmate.workers.conversion_worker.task_runner import (
    EngineUnavailableError,
    TaskRunOptions,
    execute_task,
)

# pid: (image, parent pid)
_TABLE = {
    4: ("system", 0),
    1436: ("svchost.exe", 900),
    5000: ("explorer.exe", 4),
    100: ("hwp.exe", 5000),  # 변환 전부터 사용자가 열어 둔 한글
    200: ("hwp.exe", 1436),  # COM(DcomLaunch)이 띄운 한글
    300: ("hwp.exe", 5000),  # 변환 시작 중 사용자가 연 한글
    310: ("hwpctrl.exe", 200),  # 앱 한글의 자식
    320: ("hwp.exe", 100),  # 사용자 한글의 자식
    400: ("hwp.exe", 7777),  # 부모가 이미 종료
    500: ("notepad.exe", 1436),  # 한글이 아님(PID 재사용)
}


# --- ISSUE-001 -------------------------------------------------------------


def test_select_com_launched_pids_excludes_user_launched_hwp() -> None:
    accepted, rejected = select_com_launched_pids({200, 300}, _TABLE)
    assert accepted == {200}
    assert rejected == {300}


def test_select_com_launched_pids_follows_owned_children_only() -> None:
    accepted, rejected = select_com_launched_pids({200, 310, 320}, _TABLE)
    assert accepted == {200, 310}
    assert rejected == {320}


def test_select_com_launched_pids_unknown_parent_policy() -> None:
    # 후보 하나뿐이면 기존 동작 유지(채택), 여럿이면 보수적으로 제외
    assert select_com_launched_pids({400}, _TABLE) == ({400}, set())
    assert select_com_launched_pids({200, 400}, _TABLE) == ({200}, {400})
    assert select_com_launched_pids({400}, _TABLE, allow_unknown=False) == (set(), {400})


def test_select_com_launched_pids_drops_dead_and_non_hwp_pids() -> None:
    assert select_com_launched_pids({500, 999}, _TABLE) == (set(), {500, 999})


def test_select_com_launched_pids_without_table_keeps_candidates(monkeypatch) -> None:
    import hwpmate.services.hwp_converter.process_snapshot as process_snapshot

    monkeypatch.setattr(process_snapshot, "_snapshot_process_table", lambda: None)
    assert select_com_launched_pids({200, 300}) == ({200, 300}, set())


class _DispatchHwp:
    def RegisterModule(self, *args: Any) -> bool:
        return True

    def SetMessageBoxMode(self, mode: int) -> None:
        del mode

    def Clear(self, option: int = 0) -> None:
        del option

    def Quit(self) -> None:
        return None


def _patch_initialize(monkeypatch, snapshots: list[set[int]]) -> None:
    import hwpmate.services.hwp_converter.converter as converter_impl
    import hwpmate.services.hwp_converter.process_snapshot as process_snapshot

    class FakeClient:
        @staticmethod
        def Dispatch(progid: str) -> _DispatchHwp:
            del progid
            return _DispatchHwp()

    sequence = list(snapshots)
    monkeypatch.setattr(converter_impl, "require_pywin32", lambda: (object(), FakeClient))
    monkeypatch.setattr(
        converter_impl, "ensure_hwp_security_module", lambda: (True, "ok", "FilePathCheckerModuleExample")
    )
    monkeypatch.setattr(converter_impl, "_snapshot_hwp_pids", lambda: sequence.pop(0) if sequence else set())
    monkeypatch.setattr(converter_impl, "get_snapshot_health", lambda: (0, None))
    monkeypatch.setattr(converter_impl.time, "sleep", lambda _s: None)
    monkeypatch.setattr(process_snapshot, "_snapshot_process_table", lambda: dict(_TABLE))
    monkeypatch.setattr(HWPConverter, "_suppress_hwp_ui_flash", lambda self: None)


def test_initialize_does_not_claim_hwp_opened_by_user_during_dispatch(monkeypatch) -> None:
    _patch_initialize(monkeypatch, [{100}, {100, 200, 300}])
    converter = HWPConverter()

    assert converter.initialize(manage_com_apartment=False) is True
    assert converter.owned_pids == {200}


def test_initialize_with_only_user_hwp_tracks_nothing(monkeypatch) -> None:
    _patch_initialize(monkeypatch, [{100}, {100, 300}])
    converter = HWPConverter()

    converter.initialize(manage_com_apartment=False)

    assert converter.owned_pids == set()
    assert converter.process_tracking_warning is not None


def test_mcp_live_hwp_pids_excludes_user_launched(monkeypatch) -> None:
    import hwpmate.mcp.runner as runner
    import hwpmate.services.hwp_converter.process_snapshot as process_snapshot

    monkeypatch.setattr(runner.os, "name", "nt")
    monkeypatch.setattr(process_snapshot, "_snapshot_process_table", lambda: dict(_TABLE))

    live = runner._live_hwp_pids()

    assert 200 in live and 310 in live
    assert not live & {100, 300, 320, 400}


# --- ISSUE-002 -------------------------------------------------------------


class _ComError(Exception):
    def __init__(self, hresult: int) -> None:
        super().__init__(hresult, "RPC 서버를 사용할 수 없습니다.", None, None)
        self.hresult = hresult


def test_is_com_disconnect_error_detects_rpc_failures() -> None:
    assert is_com_disconnect_error(_ComError(-2147023174))  # RPC_S_SERVER_UNAVAILABLE
    assert is_com_disconnect_error(_ComError(0x80010108))  # RPC_E_DISCONNECTED (unsigned)
    assert not is_com_disconnect_error(_ComError(-2147352567))  # DISP_E_EXCEPTION
    assert not is_com_disconnect_error(RuntimeError("SaveAs failed"))


class CrashingEngine:
    """crash_at 번째 변환부터 연결이 끊기고, initialize 하면 다시 살아나는 fake 엔진."""

    def __init__(self, *, crash_at: int, reconnect_ok: bool = True, crash_again: bool = False) -> None:
        self.crash_at = crash_at
        self.reconnect_ok = reconnect_ok
        self.crash_again = crash_again
        self.calls = 0
        self.initialize_calls = 0
        self.cleanup_calls = 0
        self.alive = True
        self.pdf_export_mode = "saveas_first"
        self.owned_pids: set[int] = {1}

    @property
    def progid_used(self) -> str | None:
        return "HWPFrame.HwpObject"

    def initialize(self, *, manage_com_apartment: bool = True) -> bool:
        del manage_com_apartment
        self.initialize_calls += 1
        if self.initialize_calls > 1 and not self.reconnect_ok:
            raise RuntimeError("Dispatch 실패")
        self.alive = True
        self.owned_pids = {self.initialize_calls}
        return True

    def convert_file(self, input_path, output_path, format_type="PDF", *, cancel_check=None):
        del input_path, output_path, format_type, cancel_check
        self.calls += 1
        if self.calls == self.crash_at or (self.crash_again and self.calls > self.crash_at):
            self.alive = False
        if not self.alive:
            return False, "RPC 서버를 사용할 수 없습니다."
        return True, None

    def is_alive(self) -> bool:
        return self.alive

    def cleanup(self) -> None:
        self.cleanup_calls += 1

    def has_owned_processes(self) -> bool:
        return bool(self.owned_pids)

    def kill_owned_processes(self) -> bool:
        return True


def _make_tasks(tmp_path: Path, count: int) -> list[ConversionTask]:
    tasks = []
    for index in range(count):
        source = tmp_path / f"doc{index}.hwp"
        source.write_bytes(b"x")
        tasks.append(ConversionTask(input_file=source, output_file=tmp_path / f"doc{index}.pdf"))
    return tasks


@pytest.fixture
def _fast_retry(monkeypatch):
    import hwpmate.workers.conversion_worker.task_runner as runner_module

    monkeypatch.setattr(runner_module, "RETRY_DELAY_SECONDS", 0)
    monkeypatch.setattr(runner_module, "RECYCLE_INITIALIZE_DELAY_SECONDS", 0)


def _run(task: ConversionTask, engine: CrashingEngine, reconnects: list[int] | None = None) -> list[str]:
    return execute_task(
        task,
        engine,
        TaskRunOptions(format_type="PDF", backup_enabled=False, retry_count=0),
        planner=TaskPlanner(),
        used_output_path_keys=set(),
        cancel_check=lambda: False,
        on_reconnected=(lambda: reconnects.append(1)) if reconnects is not None else (lambda: None),
    )


def test_execute_task_reconnects_once_and_retries_same_file(tmp_path: Path, _fast_retry) -> None:
    engine = CrashingEngine(crash_at=1)
    task = _make_tasks(tmp_path, 1)[0]
    reconnects: list[int] = []

    warnings = _run(task, engine, reconnects)

    assert task.status == "성공"
    assert engine.calls == 2
    assert engine.cleanup_calls == 1 and engine.initialize_calls == 1  # 재연결 1회
    assert reconnects == [1]
    assert any("다시 연결했습니다" in w and "RPC 서버" in w for w in warnings)


def test_execute_task_reconnects_only_once_per_file(tmp_path: Path, _fast_retry) -> None:
    engine = CrashingEngine(crash_at=1, crash_again=True)
    task = _make_tasks(tmp_path, 1)[0]

    _run(task, engine)

    assert task.status == "실패"
    assert engine.initialize_calls == 1  # 재연결은 파일당 한 번뿐


def test_execute_task_raises_when_reconnect_fails(tmp_path: Path, _fast_retry) -> None:
    engine = CrashingEngine(crash_at=1, reconnect_ok=False)
    engine.initialize_calls = 1  # 이미 한 번 초기화된 상태
    task = _make_tasks(tmp_path, 1)[0]

    with pytest.raises(EngineUnavailableError) as excinfo:
        _run(task, engine)

    assert task.status == "실패"
    assert task.error is not None and "RPC 서버" in task.error and "재연결에 실패" in task.error
    assert any("재연결에 실패" in w for w in excinfo.value.warnings)


def test_worker_recovers_after_hwp_crash_mid_batch(tmp_path: Path, _fast_retry) -> None:
    tasks = _make_tasks(tmp_path, 5)
    plan = PlannedConversion(
        format_type="PDF", same_location=True, output_path=str(tmp_path),
        backup_enabled=False, retry_count=0, tasks=tasks,
    )
    engine = CrashingEngine(crash_at=3)
    worker = ConversionWorker(plan, converter_factory=lambda: engine)
    summaries: list[Any] = []
    worker.task_completed.connect(summaries.append)

    worker.run()

    assert [task.status for task in tasks] == ["성공"] * 5
    assert engine.initialize_calls == 2
    assert any("다시 연결했습니다" in w for w in summaries[0].warnings)


def test_worker_stops_batch_when_reconnect_fails(tmp_path: Path, _fast_retry) -> None:
    tasks = _make_tasks(tmp_path, 5)
    plan = PlannedConversion(
        format_type="PDF", same_location=True, output_path=str(tmp_path),
        backup_enabled=False, retry_count=0, tasks=tasks,
    )
    engine = CrashingEngine(crash_at=3, reconnect_ok=False)
    worker = ConversionWorker(plan, converter_factory=lambda: engine)

    worker.run()

    assert [task.status for task in tasks[:2]] == ["성공", "성공"]
    assert all(task.status == "실패" for task in tasks[2:])
    assert tasks[3].error is not None and "남은 작업을 중단" in tasks[3].error
    assert engine.calls == 3


def test_is_alive_reports_dead_owned_processes(monkeypatch) -> None:
    import hwpmate.services.hwp_converter.converter as converter_impl

    converter = HWPConverter()
    converter.hwp = _DispatchHwp()  # type: ignore[assignment]
    converter.is_initialized = True
    converter.owned_pids = {200}
    monkeypatch.setattr(converter_impl, "get_snapshot_health", lambda: (0, None))

    monkeypatch.setattr(converter_impl, "_snapshot_hwp_pids", lambda: {200})
    assert converter.is_alive() is True
    monkeypatch.setattr(converter_impl, "_snapshot_hwp_pids", lambda: set())
    assert converter.is_alive() is False

    monkeypatch.setattr(converter_impl, "_snapshot_hwp_pids", lambda: {200})
    converter.connection_lost = True
    assert converter.is_alive() is False


def test_convert_file_flags_connection_lost_on_rpc_error(tmp_path: Path) -> None:
    class DeadHwp(_DispatchHwp):
        def Open(self, *args: Any) -> None:
            raise _ComError(-2147023174)

    source = tmp_path / "a.hwp"
    source.write_bytes(b"x")
    converter = HWPConverter()
    converter.hwp = DeadHwp()  # type: ignore[assignment]
    converter.is_initialized = True

    success, error = converter.convert_file(source, tmp_path / "a.pdf", "PDF")

    assert success is False and error
    assert converter.connection_lost is True
    assert converter.is_alive() is False


# --- 재순환 시 고아 한글 정리 -------------------------------------------------


def test_cleanup_terminates_previous_process_that_did_not_quit(monkeypatch) -> None:
    import hwpmate.services.hwp_converter.converter as converter_impl

    monkeypatch.setattr(converter_impl, "CLEANUP_EXIT_WAIT_SECONDS", 0)
    monkeypatch.setattr(converter_impl, "_snapshot_hwp_pids", lambda: {200, 999})
    terminated: list[set[int]] = []
    monkeypatch.setattr(
        HWPConverter, "_terminate_hwp_pids", lambda self, pids: terminated.append(set(pids)) or (True, set())
    )
    converter = HWPConverter()
    converter.hwp = _DispatchHwp()  # type: ignore[assignment]
    converter.is_initialized = True
    converter.owned_pids = {200}

    converter.cleanup()

    assert terminated == [{200}]
    assert converter.owned_pids == set()


def test_cleanup_does_not_terminate_when_process_exited(monkeypatch) -> None:
    import hwpmate.services.hwp_converter.converter as converter_impl

    monkeypatch.setattr(converter_impl, "_snapshot_hwp_pids", lambda: {999})
    terminated: list[object] = []
    monkeypatch.setattr(HWPConverter, "_terminate_hwp_pids", lambda self, pids: terminated.append(pids))
    converter = HWPConverter()
    converter.hwp = _DispatchHwp()  # type: ignore[assignment]
    converter.is_initialized = True
    converter.owned_pids = {200}

    converter.cleanup()

    assert terminated == []


# --- ISSUE-004 -------------------------------------------------------------


def test_cli_skipped_only_report_failure_is_a_warning(tmp_path: Path, capsys) -> None:
    (tmp_path / "a.hwpx").write_bytes(b"x")
    blocker = tmp_path / "blocker"
    blocker.write_text("file, not a folder", encoding="utf-8")

    ret = _run_cli_conversion(
        _parse_args(
            ["--input", str(tmp_path), "--format", "HWPX", "--report", str(blocker / "sub" / "r.json")]
        )
    )

    captured = capsys.readouterr()
    assert ret == 0
    assert "결과 저장 실패" in captured.err
    assert "Traceback" not in captured.err


# --- CLI 재순환 COM apartment (실측으로 발견) --------------------------------


def test_recycle_reopens_com_apartment_when_converter_owned_it() -> None:
    """CLI 는 initialize(manage_com_apartment=True) 로 apartment 를 컨버터가 소유한다.

    cleanup() 이 CoUninitialize 하므로 재초기화도 apartment 를 다시 열어야 한다
    (그렇지 않으면 실제 한글에서 'CoInitialize가 호출되지 않았습니다' 로 재순환·재연결 실패).
    """
    from hwpmate.workers.conversion_worker.task_runner import recycle_converter

    class ApartmentEngine(CrashingEngine):
        def __init__(self, owned: bool) -> None:
            super().__init__(crash_at=0)
            self._com_apartment_owned = owned
            self.manage_flags: list[bool] = []

        def initialize(self, *, manage_com_apartment: bool = True) -> bool:
            self.manage_flags.append(manage_com_apartment)
            return super().initialize(manage_com_apartment=manage_com_apartment)

    cli_engine = ApartmentEngine(owned=True)
    gui_engine = ApartmentEngine(owned=False)

    assert recycle_converter(cli_engine, TaskRunOptions(format_type="PDF")) is None
    assert recycle_converter(gui_engine, TaskRunOptions(format_type="PDF")) is None
    assert cli_engine.manage_flags == [True]
    assert gui_engine.manage_flags == [False]
