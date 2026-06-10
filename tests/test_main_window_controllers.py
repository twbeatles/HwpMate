from __future__ import annotations

from pathlib import Path
from typing import Any

import pytest

from hwpmate.models import AppConfig, ConversionTask, PlannedConversion


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
        def __init__(self, summary, parent=None):
            del parent
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
