from __future__ import annotations

from typing import Any, cast

from hwpmate import windows_integration
from hwpmate.windows_integration import NativeDropFilter


class FakeShell32:
    def __init__(self, paths):
        self.paths = paths
        self.finished = False

    def DragQueryFileW(self, hdrop, index, buffer, size):
        del hdrop, size
        if index == 0xFFFFFFFF:
            return len(self.paths)
        path = self.paths[index]
        if buffer is None:
            return len(path)
        buffer.value = path
        return len(path)

    def DragFinish(self, hdrop):
        del hdrop
        self.finished = True


def test_get_dropped_files_handles_paths_longer_than_max_path() -> None:
    long_path = "C:\\very-long\\" + ("nested\\" * 40) + "document.hwp"
    drop_filter = NativeDropFilter()
    fake_shell32 = FakeShell32([long_path])
    drop_filter._shell32 = cast(Any, fake_shell32)

    files = drop_filter._get_dropped_files(1)

    assert files == [long_path]
    assert len(files[0]) > 260
    assert fake_shell32.finished is True


def test_native_admin_drag_drop_policy_disabled_by_default_on_python_314(monkeypatch) -> None:
    monkeypatch.delenv(windows_integration.NATIVE_DND_DISABLE_ENV, raising=False)
    monkeypatch.delenv(windows_integration.NATIVE_DND_FORCE_ENV, raising=False)
    monkeypatch.setattr(windows_integration, "is_running_under_idle", lambda: False)
    windows_integration.get_native_admin_drag_drop_policy.cache_clear()

    enabled, reason = windows_integration.get_native_admin_drag_drop_policy()

    assert enabled is True
    assert reason == "default"


def test_native_admin_drag_drop_policy_disabled_under_idle(monkeypatch) -> None:
    monkeypatch.delenv(windows_integration.NATIVE_DND_DISABLE_ENV, raising=False)
    monkeypatch.delenv(windows_integration.NATIVE_DND_FORCE_ENV, raising=False)
    monkeypatch.setattr(windows_integration, "is_running_under_idle", lambda: True)
    windows_integration.get_native_admin_drag_drop_policy.cache_clear()

    enabled, reason = windows_integration.get_native_admin_drag_drop_policy()

    assert enabled is False
    assert "IDLE 환경" in reason


def test_native_admin_drag_drop_policy_force_env_overrides_default(monkeypatch) -> None:
    monkeypatch.setenv(windows_integration.NATIVE_DND_FORCE_ENV, "1")
    monkeypatch.setattr(windows_integration, "is_running_under_idle", lambda: True)
    windows_integration.get_native_admin_drag_drop_policy.cache_clear()

    enabled, reason = windows_integration.get_native_admin_drag_drop_policy()

    assert enabled is True
    assert windows_integration.NATIVE_DND_FORCE_ENV in reason


def test_bring_hwp_windows_to_foreground_raises_matching_hwnds(monkeypatch) -> None:
    monkeypatch.setattr(
        windows_integration,
        "_list_top_level_hwnds_for_pids",
        lambda pids, **kwargs: [11, 22] if 100 in pids else [],
    )
    raised: list[int] = []
    monkeypatch.setattr(
        windows_integration,
        "_raise_hwnd_to_foreground",
        lambda hwnd: raised.append(hwnd) or True,
    )

    count = windows_integration.bring_hwp_windows_to_foreground({100, 200})

    assert count == 2
    assert raised == [11, 22]


def test_is_likely_hwp_security_dialog_excludes_main_editor_title(monkeypatch) -> None:
    """메인 창 제목('… - 한글')은 보안 대화상자로 보지 않는다."""

    class FakeUser32:
        def IsWindow(self, hwnd):
            return True

        def GetClassNameW(self, hwnd, buf, size):
            del size
            buf.value = "HwpFrame"
            return True

        def GetWindowTextLengthW(self, hwnd):
            return len("빈 문서 1 - 한글")

        def GetWindowTextW(self, hwnd, buf, size):
            del size
            buf.value = "빈 문서 1 - 한글"
            return True

    import ctypes as ct

    class Windll:
        user32 = FakeUser32()

    monkeypatch.setattr(ct, "windll", Windll())
    assert windows_integration.is_likely_hwp_security_dialog(1) is False


def test_is_likely_hwp_security_dialog_matches_security_title(monkeypatch) -> None:
    class FakeUser32:
        def IsWindow(self, hwnd):
            return True

        def GetClassNameW(self, hwnd, buf, size):
            del size
            buf.value = "SomeClass"
            return True

        def GetWindowTextLengthW(self, hwnd):
            return len("보안 경고")

        def GetWindowTextW(self, hwnd, buf, size):
            del size
            buf.value = "보안 경고"
            return True

    import ctypes as ct

    class Windll:
        user32 = FakeUser32()

    monkeypatch.setattr(ct, "windll", Windll())
    assert windows_integration.is_likely_hwp_security_dialog(1) is True


def test_hide_hwp_main_windows_skips_security_dialogs(monkeypatch) -> None:
    monkeypatch.setattr(
        windows_integration,
        "_resolve_hwp_target_pids",
        lambda pids=None: {10},
    )
    monkeypatch.setattr(
        windows_integration,
        "_list_top_level_hwnds_for_pids",
        lambda pids, **kwargs: [1, 2],
    )
    monkeypatch.setattr(
        windows_integration,
        "is_likely_hwp_security_dialog",
        lambda hwnd: hwnd == 2,
    )

    class FakeUser32:
        def __init__(self) -> None:
            self.hide_calls: list[int] = []
            self._visible = {1: True, 2: True}

        def IsWindowVisible(self, hwnd):
            return self._visible.get(hwnd, False)

        def ShowWindow(self, hwnd, cmd):
            if cmd == 0:
                self.hide_calls.append(hwnd)
                self._visible[hwnd] = False
            return 1

    import ctypes as ct

    fake = FakeUser32()

    class Windll:
        user32 = fake

    monkeypatch.setattr(ct, "windll", Windll())

    hidden = windows_integration.hide_hwp_main_windows({10})
    assert hidden == 1
    assert fake.hide_calls == [1]


def test_suppress_hwp_ui_flash_hides_then_raises(monkeypatch) -> None:
    order: list[str] = []

    monkeypatch.setattr(
        windows_integration,
        "hide_hwp_main_windows",
        lambda pids=None: order.append("hide") or 3,
    )
    monkeypatch.setattr(
        windows_integration,
        "bring_hwp_windows_to_foreground",
        lambda pids=None: order.append("raise") or 1,
    )

    hidden, raised = windows_integration.suppress_hwp_ui_flash({1})
    assert (hidden, raised) == (3, 1)
    assert order == ["hide", "raise"]


def test_bring_hwp_windows_to_foreground_returns_zero_without_pids(monkeypatch) -> None:
    monkeypatch.setattr(
        "hwpmate.services.hwp_converter._snapshot_hwp_pids",
        lambda: set(),
    )
    count = windows_integration.bring_hwp_windows_to_foreground(None)
    assert count == 0


def test_try_accept_hwp_security_dialog_clicks_matching_button(monkeypatch) -> None:
    monkeypatch.setattr(
        windows_integration,
        "_resolve_hwp_target_pids",
        lambda pids=None: {1},
    )
    monkeypatch.setattr(
        windows_integration,
        "_list_top_level_hwnds_for_pids",
        lambda pids, **kwargs: [100],
    )
    monkeypatch.setattr(windows_integration, "_raise_hwnd_to_foreground", lambda hwnd: True)

    class FakeUser32:
        def EnumChildWindows(self, parent, callback, lparam):
            del parent, lparam
            # 자식 버튼 시뮬레이션
            callback(200, 0)

        def GetWindowTextLengthW(self, hwnd):
            return 4 if hwnd == 200 else 0

        def GetWindowTextW(self, hwnd, buf, size):
            del size
            if hwnd == 200:
                buf.value = "모두 허용"
            return True

        def GetClassNameW(self, hwnd, buf, size):
            del size
            if hwnd == 200:
                buf.value = "Button"
            return True

        def SendMessageW(self, hwnd, msg, wparam, lparam):
            self.last = (hwnd, msg, wparam, lparam)
            return 0

    fake = FakeUser32()
    monkeypatch.setattr(windows_integration.ctypes.windll, "user32", fake, raising=False)
    # windll.user32 is not always patchable this way; patch module-level access via lambda in function
    import ctypes as ct

    class Windll:
        user32 = fake

    monkeypatch.setattr(ct, "windll", Windll())

    assert windows_integration.try_accept_hwp_security_dialog({1}) is True
    assert fake.last[0] == 200
    assert fake.last[1] == windows_integration._BM_CLICK
