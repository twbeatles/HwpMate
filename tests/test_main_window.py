from __future__ import annotations

from pathlib import Path

from PyQt6.QtGui import QCloseEvent
from PyQt6.QtWidgets import QMessageBox

from hwpmate.models import AppConfig


class DummyTray:
    def hide(self) -> None:
        return None


class FakeWorker:
    def __init__(self) -> None:
        self.cancel_called = False

    def cancel(self) -> None:
        self.cancel_called = True

    def wait(self, timeout: int) -> bool:
        del timeout
        return False

    def can_force_terminate(self) -> bool:
        return False

    def isRunning(self) -> bool:
        return True


def create_window(monkeypatch, qapp):
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
    window = main_window_module.MainWindow()
    return window, saved_configs


def test_close_event_saves_current_ui_settings(monkeypatch, qapp, tmp_path: Path) -> None:
    window, saved_configs = create_window(monkeypatch, qapp)
    window.files_radio.setChecked(True)
    window.same_location_check.setChecked(False)
    window.overwrite_check.setChecked(True)
    window.backup_check.setChecked(False)
    window.retry_spin.setValue(2)
    window.include_sub_check.setChecked(False)
    window.output_entry.setText(str(tmp_path))

    event = QCloseEvent()
    window.closeEvent(event)

    assert event.isAccepted()
    assert saved_configs
    latest = saved_configs[-1]
    assert latest["mode"] == "files"
    assert latest["same_location"] is False
    assert latest["overwrite"] is True
    assert latest["backup_enabled"] is False
    assert latest["retry_count"] == 2
    assert latest["include_sub"] is False
    assert latest["output_path"] == str(tmp_path)


def test_folder_mode_native_drop_requires_single_folder(monkeypatch, qapp, tmp_path: Path) -> None:
    window, _ = create_window(monkeypatch, qapp)
    folder = tmp_path / "folder"
    folder.mkdir()
    file_path = tmp_path / "doc.hwp"
    file_path.write_text("x", encoding="utf-8")
    warnings = []
    preview_calls = []

    monkeypatch.setattr(QMessageBox, "warning", lambda *args, **kwargs: warnings.append((args, kwargs)))
    monkeypatch.setattr(window, "_start_folder_preview_scan", lambda folder_path: preview_calls.append(folder_path))

    window.folder_radio.setChecked(True)
    window._on_native_files_dropped([str(file_path)])
    window._on_native_files_dropped([str(folder)])

    assert warnings
    assert preview_calls == [str(folder.resolve())]


def test_format_change_restarts_folder_preview(monkeypatch, qapp, tmp_path: Path) -> None:
    window, _ = create_window(monkeypatch, qapp)
    folder = tmp_path / "docs"
    folder.mkdir()
    calls = []

    monkeypatch.setattr(window, "_start_folder_preview_scan", lambda folder_path: calls.append(folder_path))

    window.folder_radio.setChecked(True)
    window.folder_entry.setText(str(folder))
    window._on_format_card_clicked("DOCX")

    assert calls == [str(folder)]


def test_close_event_waits_for_running_worker(monkeypatch, qapp) -> None:
    window, _ = create_window(monkeypatch, qapp)
    window.is_converting = True
    window.worker = FakeWorker()  # type: ignore[assignment]

    monkeypatch.setattr(QMessageBox, "question", lambda *args, **kwargs: QMessageBox.StandardButton.Yes)

    event = QCloseEvent()
    window.closeEvent(event)

    assert not event.isAccepted()
    assert window._close_after_worker is True
    assert window.worker.cancel_called is True  # type: ignore[union-attr]


def test_set_converting_state_keeps_output_button_disabled_for_same_location(monkeypatch, qapp) -> None:
    window, _ = create_window(monkeypatch, qapp)
    window.same_location_check.setChecked(True)

    window._set_converting_state(True)
    window._set_converting_state(False)

    assert window.output_btn.isEnabled() is False
    assert window.output_entry.isEnabled() is False


def test_start_conversion_shows_result_for_skipped_only_plan(monkeypatch, qapp, tmp_path: Path) -> None:
    window, _ = create_window(monkeypatch, qapp)
    skipped = tmp_path / "same.hwpx"
    skipped.write_text("x", encoding="utf-8")
    window.files_radio.setChecked(True)
    window._selected_format = "HWPX"
    window.file_store.add_paths([str(skipped)])
    shown = []

    class FakeResultDialog:
        def __init__(self, summary, parent=None):
            del parent
            shown.append(summary)

        def exec(self):
            return None

    import hwpmate.ui.main_window as main_window_module

    monkeypatch.setattr(main_window_module, "ResultDialog", FakeResultDialog)

    window._start_conversion()

    assert len(shown) == 1
    assert shown[0].skipped_count == 1
    assert window.worker is None
