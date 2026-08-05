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
            # 전체 확장자 캐시가 있으면 재스캔 없이 변환 가능 수만 갱신한다.
            if self.state.folder_scan_ready:
                self.window.file_selection_controller.refresh_folder_preview_count()
            else:
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
        can_select_output = (not same_location) and (
            not self.state.is_converting and not self.state.is_planning
        )
        self.window.output_entry.setEnabled(can_select_output)
        self.window.output_btn.setEnabled(can_select_output)

    def on_include_sub_toggled(self, _: bool) -> None:
        if self.window.folder_radio.isChecked() and self.window.folder_entry.text().strip():
            self.window.file_selection_controller.invalidate_folder_scan_cache()
            self.window._start_folder_preview_scan(self.window.folder_entry.text().strip())

    def set_converting_state(self, converting: bool) -> None:
        if converting:
            self.window._cancel_active_scan()
            # 변환 시작 시 계획 잠금은 해제 (converting 이 잠금 역할)
            self.state.is_planning = False

        self.state.is_converting = converting
        if not converting:
            self.state.force_kill_pending = False
            self.window.cancel_btn.setText("⏹️ 취소")

        self.apply_busy_ui()

    def set_planning_state(self, planning: bool) -> None:
        """스캔 대기·작업 수집 중 입력/시작 재진입을 막는다."""
        self.state.is_planning = planning
        self.apply_busy_ui()

    def apply_busy_ui(self) -> None:
        """변환 중 또는 계획 중일 때 입력·시작을 잠근다."""
        converting = self.state.is_converting
        planning = self.state.is_planning
        busy = converting or planning

        self.window.start_btn.setEnabled(not busy)
        # 취소는 실제 변환 중에만
        self.window.cancel_btn.setEnabled(converting)
        if hasattr(self.window, "lifecycle_controller"):
            self.window.lifecycle_controller.set_command_actions_enabled(not busy)

        self.window.folder_radio.setEnabled(not busy)
        self.window.files_radio.setEnabled(not busy)

        for card in self.window.format_cards.values():
            card.setEnabled(not busy)

        self.window.same_location_check.setEnabled(not busy)
        self.window.overwrite_check.setEnabled(not busy)
        self.window.backup_check.setEnabled(not busy)
        backup_max = getattr(self.window, "backup_max_spin", None)
        if backup_max is not None:
            backup_max.setEnabled(not busy)
        auto_accept = getattr(self.window, "auto_accept_security_check", None)
        if auto_accept is not None:
            auto_accept.setEnabled(not busy)
        pdf_combo = getattr(self.window, "pdf_export_mode_combo", None)
        if pdf_combo is not None:
            pdf_combo.setEnabled(not busy)
        self.window.retry_spin.setEnabled(not busy)
        self.window.include_sub_check.setEnabled(not busy)

        if hasattr(self.window, "drop_area"):
            self.window.drop_area.setEnabled(not busy)
        if hasattr(self.window, "add_btn"):
            self.window.add_btn.setEnabled(not busy)
        if hasattr(self.window, "remove_btn"):
            self.window.remove_btn.setEnabled(not busy)
        if hasattr(self.window, "clear_btn"):
            self.window.clear_btn.setEnabled(not busy)
        if hasattr(self.window, "folder_btn"):
            self.window.folder_btn.setEnabled(not busy)
        if hasattr(self.window, "theme_btn"):
            self.window.theme_btn.setEnabled(not busy)

        self.update_output_ui()
