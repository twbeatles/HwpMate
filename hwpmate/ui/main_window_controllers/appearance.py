from __future__ import annotations

from typing import Any, Callable

from ...constants import FORMAT_TYPES
from ...models import AppConfig
from ..theme import ThemeManager
from .state import MainWindowState


class AppearanceController:
    """Theme, format selection, and enabled-state orchestration."""

    def __init__(
        self,
        window: Any,
        state: MainWindowState,
        save_config_func: Callable[[AppConfig], bool],
    ) -> None:
        self.window = window
        self.state = state
        self._save_config = save_config_func

    def apply_theme(self) -> None:
        theme_css = ThemeManager.get_theme(self.window.current_theme)
        self.window.setStyleSheet(theme_css)

    def toggle_theme(self) -> None:
        if self.window.current_theme == "dark":
            self.window.current_theme = "light"
            self.window.theme_btn.setText("☀️ 라이트")
        else:
            self.window.current_theme = "dark"
            self.window.theme_btn.setText("🌙 다크")

        self.apply_theme()
        self.window.config["theme"] = self.window.current_theme
        if self._save_config(self.window.config) is False:
            self.window.status_label.setText("테마 설정 저장에 실패했습니다")
            if hasattr(self.window, "toast"):
                self.window.toast.show_message("테마 설정 저장에 실패했습니다", "⚠️")

    def on_format_card_clicked(self, format_type: str) -> None:
        if format_type not in FORMAT_TYPES:
            return

        self.state.selected_format = format_type
        self.update_format_cards()
        if self.window.folder_radio.isChecked() and self.window.folder_entry.text().strip():
            self.window._start_folder_preview_scan(self.window.folder_entry.text().strip())

    def update_format_cards(self) -> None:
        for fmt_key, card in self.window.format_cards.items():
            card.setSelected(self.state.selected_format == fmt_key)

    def update_mode_ui(self, *_: object) -> None:
        self.window._cancel_active_scan()
        is_folder_mode = self.window.folder_radio.isChecked()
        self.window.folder_widget.setVisible(is_folder_mode)
        self.window.files_widget.setVisible(not is_folder_mode)

    def update_output_ui(self, *_: object) -> None:
        same_location = self.window.same_location_check.isChecked()
        can_select_output = (not same_location) and (not self.state.is_converting)
        self.window.output_entry.setEnabled(can_select_output)
        self.window.output_btn.setEnabled(can_select_output)

    def on_include_sub_toggled(self, _: bool) -> None:
        if self.window.folder_radio.isChecked() and self.window.folder_entry.text().strip():
            self.window._start_folder_preview_scan(self.window.folder_entry.text().strip())

    def set_converting_state(self, converting: bool) -> None:
        if converting:
            self.window._cancel_active_scan()

        self.state.is_converting = converting
        self.window.start_btn.setEnabled(not converting)
        self.window.cancel_btn.setEnabled(converting)
        if hasattr(self.window, "lifecycle_controller"):
            self.window.lifecycle_controller.set_command_actions_enabled(not converting)

        if not converting:
            self.state.force_kill_pending = False
            self.window.cancel_btn.setText("⏹️ 취소")

        self.window.folder_radio.setEnabled(not converting)
        self.window.files_radio.setEnabled(not converting)

        for card in self.window.format_cards.values():
            card.setEnabled(not converting)

        self.window.same_location_check.setEnabled(not converting)
        self.window.overwrite_check.setEnabled(not converting)
        self.window.backup_check.setEnabled(not converting)
        self.window.retry_spin.setEnabled(not converting)
        self.window.include_sub_check.setEnabled(not converting)

        if hasattr(self.window, "drop_area"):
            self.window.drop_area.setEnabled(not converting)
        if hasattr(self.window, "add_btn"):
            self.window.add_btn.setEnabled(not converting)
        if hasattr(self.window, "remove_btn"):
            self.window.remove_btn.setEnabled(not converting)
        if hasattr(self.window, "clear_btn"):
            self.window.clear_btn.setEnabled(not converting)
        if hasattr(self.window, "folder_btn"):
            self.window.folder_btn.setEnabled(not converting)

        self.update_output_ui()
