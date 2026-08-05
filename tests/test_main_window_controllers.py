from __future__ import annotations

from pathlib import Path
from typing import Any

import pytest

from hwpmate.models import AppConfig, ConversionSummary, ConversionTask, PlannedConversion


class DummyTray:
    def hide(self) -> None:
        return None


class FakeSignal:
    def __init__(self) -> None:
        self.disconnect_count = 0

    def disconnect(self, *_: object) -> None:
        self.disconnect_count += 1


class FakeScanWorker:
    def __init__(self) -> None:
        self.batch_found = FakeSignal()
        self.scan_progress = FakeSignal()
        self.scan_finished = FakeSignal()
        self.scan_error = FakeSignal()
        self.finished = FakeSignal()
        self.deleted = False

    def isRunning(self) -> bool:
        return False

    def deleteLater(self) -> None:
        self.deleted = True


def create_window(monkeypatch: pytest.MonkeyPatch, qapp: Any):
    del qapp
    import hwpmate.ui.main_window as main_window_module

    saved_configs = []
    monkeypatch.setattr(main_window_module, "load_config", lambda: AppConfig())
    monkeypatch.setattr(
        main_window_module,
        "save_config",
        lambda config: saved_configs.append(config.to_dict() if hasattr(config, "to_dict") else dict(config)),
    )
    monkeypatch.setattr(
        main_window_module.MainWindow,
        "_init_tray_icon",
        lambda self: setattr(self, "tray_icon", DummyTray()),
    )
    return main_window_module.MainWindow(), saved_configs


def test_file_selection_controller_clears_finished_scan_worker(monkeypatch: pytest.MonkeyPatch, qapp: Any) -> None:
    window, _ = create_window(monkeypatch, qapp)
    fake_worker = FakeScanWorker()
    window.state.scan_worker = fake_worker  # type: ignore[assignment]
    window.state.scan_mode = "add_files"
    window.state.scan_new_file_count = 3
    window.state.scan_preview_count = 5

    assert window.file_selection_controller.cancel_active_scan() is True

    assert fake_worker.deleted is True
    assert window.file_scan_worker is None
    assert window._scan_mode is None
    assert window._scan_new_file_count == 0
    assert window._scan_preview_count == 0


def test_file_selection_controller_appends_unique_files(monkeypatch: pytest.MonkeyPatch, qapp: Any, tmp_path: Path) -> None:
    window, _ = create_window(monkeypatch, qapp)
    first = tmp_path / "a.hwp"
    second = tmp_path / "b.hwpx"
    first.write_text("x", encoding="utf-8")
    second.write_text("x", encoding="utf-8")

    added = window.file_selection_controller.append_files_batch([str(first), str(first), str(second)])

    assert added == 2
    assert window.file_store.count == 2
    assert window.file_table.rowCount() == 2
    first_item = window.file_table.item(0, 0)
    assert first_item is not None
    assert first_item.text() == "a.hwp"
    assert window.file_count_label.text() == "📄 파일: 2개"


def test_conversion_controller_validates_custom_output_folder(monkeypatch: pytest.MonkeyPatch, qapp: Any) -> None:
    window, _ = create_window(monkeypatch, qapp)
    window.same_location_check.setChecked(False)
    window.output_entry.setText("")

    with pytest.raises(ValueError, match="출력 폴더"):
        window.conversion_controller.validate_output_settings()


def test_conversion_controller_shows_skipped_only_result(monkeypatch: pytest.MonkeyPatch, qapp: Any, tmp_path: Path) -> None:
    window, _ = create_window(monkeypatch, qapp)
    skipped_file = tmp_path / "same.hwpx"
    skipped_file.write_text("x", encoding="utf-8")
    plan = PlannedConversion(
        format_type="HWPX",
        same_location=True,
        output_path="",
        skipped_tasks=[
            ConversionTask(skipped_file, skipped_file, status="건너뜀", error="이미 HWPX 형식입니다."),
        ],
    )
    shown = []

    class FakeResultDialog:
        def __init__(self, summary, parent=None, **kwargs):
            del parent, kwargs
            shown.append(summary)

        def exec(self):
            return None

    import hwpmate.ui.main_window as main_window_module

    monkeypatch.setattr(main_window_module, "ResultDialog", FakeResultDialog)

    window.conversion_controller.show_skipped_only_result(plan)

    assert len(shown) == 1
    assert shown[0].skipped_count == 1
    assert window.last_summary is shown[0]
    assert window.plan is None


def test_start_conversion_ignores_duplicate_start_while_converting(monkeypatch: pytest.MonkeyPatch, qapp: Any) -> None:
    window, _ = create_window(monkeypatch, qapp)
    window.is_converting = True

    window.conversion_controller.start_conversion()

    assert window.worker is None
    assert "이미 진행" in window.status_label.text()


def test_set_converting_state_disables_menu_actions_and_start_shortcut(monkeypatch: pytest.MonkeyPatch, qapp: Any) -> None:
    window, _ = create_window(monkeypatch, qapp)

    window._set_converting_state(True)

    assert window.lifecycle_controller.add_files_action is not None
    assert window.lifecycle_controller.add_files_action.isEnabled() is False
    assert window.lifecycle_controller.add_folder_action is not None
    assert window.lifecycle_controller.add_folder_action.isEnabled() is False
    assert window.lifecycle_controller.start_shortcut is not None
    assert window.lifecycle_controller.start_shortcut.isEnabled() is False

    window._set_converting_state(False)
    assert window.lifecycle_controller.add_files_action.isEnabled() is True
    assert window.lifecycle_controller.start_shortcut.isEnabled() is True


def test_file_selection_ignores_changes_while_converting(monkeypatch: pytest.MonkeyPatch, qapp: Any, tmp_path: Path) -> None:
    window, _ = create_window(monkeypatch, qapp)
    dropped = tmp_path / "doc.hwp"
    dropped.write_text("x", encoding="utf-8")
    window.is_converting = True

    window.file_selection_controller.add_files([str(dropped)])

    assert window.file_store.count == 0
    assert "변환 중" in window.status_label.text()


def test_worker_finished_preserves_failed_hwp_status(monkeypatch: pytest.MonkeyPatch, qapp: Any, tmp_path: Path) -> None:
    window, _ = create_window(monkeypatch, qapp)
    failed = ConversionTask(tmp_path / "a.hwp", tmp_path / "a.pdf", status="실패", error="boom")
    window.last_summary = ConversionSummary(
        format_type="PDF",
        tasks=[failed],
    )

    window.conversion_controller.on_worker_finished()

    assert "실패" in window.hwp_status_label.text()


def test_on_status_updated_syncs_hwp_status_label(monkeypatch: pytest.MonkeyPatch, qapp: Any) -> None:
    window, _ = create_window(monkeypatch, qapp)
    controller = window.conversion_controller

    controller.on_status_updated("한글 프로그램 연결 중... 허용/보안 창이 뜨면 확인")
    assert "허용 창" in window.hwp_status_label.text()
    assert "연결 중" in window.status_label.text()

    controller.on_status_updated("연결 성공: HWPFrame.HwpObject")
    assert "연결됨" in window.hwp_status_label.text()
    # 연결 성공 후에도 Open 구간 폴링을 유지해야 함 (중지하지 않음)
    # 타이머가 없으면 그대로 None 허용

    controller.on_status_updated("변환 중: sample.hwp")
    assert "변환 중" in window.hwp_status_label.text()


def test_retry_failed_tasks_builds_plan_and_starts_worker(
    monkeypatch: pytest.MonkeyPatch, qapp: Any, tmp_path: Path
) -> None:
    window, _ = create_window(monkeypatch, qapp)
    failed_path = tmp_path / "a.hwp"
    failed_path.write_text("x", encoding="utf-8")
    failed = ConversionTask(failed_path, failed_path.with_suffix(".pdf"), status="실패", error="boom")
    window.last_summary = ConversionSummary(format_type="PDF", tasks=[failed])
    started: list[object] = []

    class AcceptPreflight:
        def __init__(self, plan, parent=None):
            del plan, parent

        def exec(self):
            return window.dialog_accepted_code()

    class FakeWorker:
        def __init__(self, plan):
            self.plan = plan
            self.progress_updated = type("S", (), {"connect": lambda *a, **k: None})()
            self.status_updated = type("S", (), {"connect": lambda *a, **k: None})()
            self.task_completed = type("S", (), {"connect": lambda *a, **k: None})()
            self.finished = type("S", (), {"connect": lambda *a, **k: None})()
            self.engine_status_updated = type("S", (), {"connect": lambda *a, **k: None})()

        def start(self):
            started.append(self.plan)

    monkeypatch.setattr(window, "_create_preflight_dialog", lambda plan: AcceptPreflight(plan))
    monkeypatch.setattr(window, "_create_conversion_worker", lambda plan: FakeWorker(plan))
    monkeypatch.setattr(window.toast, "show_message", lambda *a, **k: None)

    window.conversion_controller.retry_failed_tasks([failed])

    assert len(started) == 1
    assert started[0].runnable_count == 1  # type: ignore[attr-defined]
    assert window.is_converting is True


def test_collect_tasks_requires_folder_cache_after_wait(
    monkeypatch: pytest.MonkeyPatch, qapp: Any
) -> None:
    window, _ = create_window(monkeypatch, qapp)
    window.folder_radio.setChecked(True)
    window.folder_entry.setText("C:/docs")

    monkeypatch.setattr(
        window.file_selection_controller,
        "get_folder_scan_cache",
        lambda **kwargs: None,
    )

    with pytest.raises(ValueError, match="폴더 스캔 결과"):
        window.conversion_controller.collect_tasks(require_folder_cache=True)


def test_poll_uses_owned_pids_and_skips_auto_accept_when_module_ok(
    monkeypatch: pytest.MonkeyPatch, qapp: Any
) -> None:
    window, _ = create_window(monkeypatch, qapp)
    controller = window.conversion_controller
    window.state.is_converting = True
    session = window.state.security_session
    session.owned_pids = {42}
    session.security_module_registered = True
    session.engine_status_received = True
    session.auto_accept_enabled = True

    brought: list[object] = []
    hidden: list[object] = []
    accepted: list[object] = []

    import hwpmate.ui.main_window_controllers.conversion as conv_mod

    monkeypatch.setattr(
        conv_mod,
        "hide_hwp_main_windows",
        lambda pids: hidden.append(pids) or 1,
    )
    monkeypatch.setattr(
        conv_mod,
        "bring_hwp_windows_to_foreground",
        lambda pids: brought.append(pids) or 1,
    )
    monkeypatch.setattr(
        conv_mod,
        "try_accept_hwp_security_dialog",
        lambda pids: accepted.append(pids) or True,
    )

    controller._poll_hwp_foreground()

    assert hidden == [{42}]
    assert brought == [{42}]
    assert accepted == []


def test_poll_auto_accept_when_module_failed(
    monkeypatch: pytest.MonkeyPatch, qapp: Any
) -> None:
    window, _ = create_window(monkeypatch, qapp)
    controller = window.conversion_controller
    window.state.is_converting = True
    session = window.state.security_session
    session.owned_pids = {7}
    session.security_module_registered = False
    session.engine_status_received = True
    session.auto_accept_enabled = True

    accepted: list[object] = []
    hidden: list[object] = []

    import hwpmate.ui.main_window_controllers.conversion as conv_mod

    monkeypatch.setattr(
        conv_mod,
        "hide_hwp_main_windows",
        lambda pids: hidden.append(pids) or 0,
    )
    monkeypatch.setattr(conv_mod, "bring_hwp_windows_to_foreground", lambda pids: 0)
    monkeypatch.setattr(
        conv_mod,
        "try_accept_hwp_security_dialog",
        lambda pids: accepted.append(pids) or True,
    )

    controller._poll_hwp_foreground()

    assert hidden == [{7}]
    assert accepted == [{7}]
    assert session.auto_accept_clicks == 1


def test_poll_skips_auto_accept_before_engine_status(
    monkeypatch: pytest.MonkeyPatch, qapp: Any
) -> None:
    window, _ = create_window(monkeypatch, qapp)
    controller = window.conversion_controller
    window.state.is_converting = True
    session = window.state.security_session
    session.owned_pids = set()
    session.security_module_registered = None
    session.engine_status_received = False
    session.auto_accept_enabled = True

    accepted: list[object] = []
    brought: list[object] = []
    hidden: list[object] = []
    import hwpmate.ui.main_window_controllers.conversion as conv_mod

    monkeypatch.setattr(
        conv_mod,
        "hide_hwp_main_windows",
        lambda pids: hidden.append(pids) or 0,
    )
    monkeypatch.setattr(
        conv_mod,
        "bring_hwp_windows_to_foreground",
        lambda pids: brought.append(pids) or 0,
    )
    monkeypatch.setattr(
        conv_mod,
        "try_accept_hwp_security_dialog",
        lambda pids: accepted.append(pids) or True,
    )

    controller._poll_hwp_foreground()
    assert accepted == []
    # engine_status 전 전면화·메인 숨김도 생략
    assert brought == []
    assert hidden == []


def test_start_conversion_blocked_while_planning(
    monkeypatch: pytest.MonkeyPatch, qapp: Any
) -> None:
    window, _ = create_window(monkeypatch, qapp)
    window.state.is_planning = True
    window.conversion_controller.start_conversion()
    assert window.worker is None
    assert "준비" in window.status_label.text() or "진행" in window.status_label.text()


def test_on_engine_status_updates_session(monkeypatch: pytest.MonkeyPatch, qapp: Any) -> None:
    window, _ = create_window(monkeypatch, qapp)
    controller = window.conversion_controller
    controller.on_engine_status_updated(
        {
            "security_module_registered": True,
            "owned_pids": [99],
            "snapshot_unreliable": False,
        }
    )
    assert window.state.security_session.owned_pids == {99}
    assert window.state.security_session.security_module_registered is True
    assert window.state.security_session.engine_status_received is True


def test_folder_cache_freshness_detects_missing_files(
    monkeypatch: pytest.MonkeyPatch, qapp: Any, tmp_path: Path
) -> None:
    window, _ = create_window(monkeypatch, qapp)
    existing = tmp_path / "a.hwp"
    existing.write_text("x", encoding="utf-8")
    missing = tmp_path / "gone.hwp"
    paths = [str(existing), str(missing), str(tmp_path / "b.hwp"), str(tmp_path / "c.hwp")]
    ok, reason = window.file_selection_controller.validate_folder_scan_cache_freshness(paths)
    assert ok is False
    assert "변경" in reason or "없음" in reason


def test_folder_cache_mtime_only_change_does_not_fail(
    monkeypatch: pytest.MonkeyPatch, qapp: Any, tmp_path: Path
) -> None:
    """디렉터리 mtime 만 바뀌고 샘플 파일이 있으면 하드 실패하지 않는다."""
    window, _ = create_window(monkeypatch, qapp)
    f1 = tmp_path / "a.hwp"
    f1.write_text("x", encoding="utf-8")
    window.state.folder_scan_dir_mtime = 1.0
    window.state.folder_scan_folder = str(tmp_path)
    window.state.folder_scan_file_count = 1
    ok, reason = window.file_selection_controller.validate_folder_scan_cache_freshness(
        [str(f1)],
        folder_path=str(tmp_path),
    )
    assert ok is True
    assert reason == ""


def test_poll_skips_window_control_without_owned_pids(
    monkeypatch: pytest.MonkeyPatch, qapp: Any
) -> None:
    window, _ = create_window(monkeypatch, qapp)
    controller = window.conversion_controller
    window.state.is_converting = True
    session = window.state.security_session
    session.owned_pids = set()
    session.security_module_registered = False
    session.engine_status_received = True
    session.auto_accept_enabled = True

    hidden: list[object] = []
    brought: list[object] = []
    accepted: list[object] = []
    import hwpmate.ui.main_window_controllers.conversion as conv_mod

    monkeypatch.setattr(
        conv_mod, "hide_hwp_main_windows", lambda pids: hidden.append(pids) or 0
    )
    monkeypatch.setattr(
        conv_mod, "bring_hwp_windows_to_foreground", lambda pids: brought.append(pids) or 0
    )
    monkeypatch.setattr(
        conv_mod, "try_accept_hwp_security_dialog", lambda pids: accepted.append(pids) or True
    )

    controller._poll_hwp_foreground()
    assert hidden == []
    assert brought == []
    assert accepted == []


def test_retry_failed_tasks_blocked_while_planning(
    monkeypatch: pytest.MonkeyPatch, qapp: Any, tmp_path: Path
) -> None:
    window, _ = create_window(monkeypatch, qapp)
    window.state.is_planning = True
    src = tmp_path / "a.hwp"
    src.write_text("x", encoding="utf-8")
    failed = [ConversionTask(input_file=src, output_file=tmp_path / "a.pdf", status="실패")]
    window.conversion_controller.retry_failed_tasks(failed)
    assert window.worker is None


def test_request_worker_stop_sets_force_kill_pending(
    monkeypatch: pytest.MonkeyPatch, qapp: Any
) -> None:
    window, _ = create_window(monkeypatch, qapp)
    controller = window.conversion_controller

    class FakeWorker:
        def __init__(self) -> None:
            self.cancel_called = False
            self._running = True

        def cancel(self) -> None:
            self.cancel_called = True

        def isRunning(self) -> bool:
            return self._running

        def wait(self, _ms: int) -> bool:
            return False

        def can_force_terminate(self) -> bool:
            return True

    fake = FakeWorker()
    window.state.worker = fake  # type: ignore[assignment]
    monkeypatch.setattr(
        "hwpmate.ui.main_window_controllers.conversion.controller.WORKER_WAIT_TIMEOUT",
        50,
    )
    ok = controller.request_worker_stop("취소 요청 중...")
    assert ok is False
    assert fake.cancel_called is True
    assert window.state.force_kill_pending is True
    assert "강제 종료" in window.cancel_btn.text() or "응답" in window.status_label.text()


def test_native_drop_controller_routes_file_mode_paths(monkeypatch: pytest.MonkeyPatch, qapp: Any, tmp_path: Path) -> None:
    window, _ = create_window(monkeypatch, qapp)
    dropped = tmp_path / "doc.hwp"
    dropped.write_text("x", encoding="utf-8")
    calls = []

    window.files_radio.setChecked(True)
    monkeypatch.setattr(window, "_add_files", lambda files: calls.append(files))
    monkeypatch.setattr(window.toast, "show_message", lambda *args, **kwargs: None)

    window.native_drop_controller.on_native_files_dropped([str(dropped)])

    assert calls == [[str(dropped.resolve())]]


def test_native_drop_ignores_paths_while_converting(monkeypatch: pytest.MonkeyPatch, qapp: Any, tmp_path: Path) -> None:
    window, _ = create_window(monkeypatch, qapp)
    dropped = tmp_path / "doc.hwp"
    dropped.write_text("x", encoding="utf-8")
    calls = []

    window.files_radio.setChecked(True)
    window.is_converting = True
    monkeypatch.setattr(window, "_add_files", lambda files: calls.append(files))
    monkeypatch.setattr(window.toast, "show_message", lambda *args, **kwargs: None)

    window.native_drop_controller.on_native_files_dropped([str(dropped)])

    assert calls == []
    assert "변환 중" in window.status_label.text()


def test_native_drop_ignores_paths_while_planning(
    monkeypatch: pytest.MonkeyPatch, qapp: Any, tmp_path: Path
) -> None:
    window, _ = create_window(monkeypatch, qapp)
    dropped = tmp_path / "doc.hwp"
    dropped.write_text("x", encoding="utf-8")
    calls = []

    window.files_radio.setChecked(True)
    window.state.is_planning = True
    monkeypatch.setattr(window, "_add_files", lambda files: calls.append(files))
    monkeypatch.setattr(window.toast, "show_message", lambda *args, **kwargs: None)

    window.native_drop_controller.on_native_files_dropped([str(dropped)])

    assert calls == []
    assert "준비" in window.status_label.text()


def test_folder_cache_sample_includes_head_and_tail() -> None:
    from hwpmate.ui.main_window_controllers.file_selection import FileSelectionController

    paths = [f"p{i}.hwp" for i in range(100)]
    sample = FileSelectionController._sample_cache_paths(paths, 24)
    assert paths[0] in sample
    assert paths[-1] in sample
    assert len(sample) == 24
    assert len(set(sample)) == 24


def test_wait_for_active_scan_returns_false_when_close_requested(
    monkeypatch: pytest.MonkeyPatch, qapp: Any
) -> None:
    window, _ = create_window(monkeypatch, qapp)
    window.state.close_requested = True
    window.state.scan_worker = None
    assert window.file_selection_controller.wait_for_active_scan(1000) is False
