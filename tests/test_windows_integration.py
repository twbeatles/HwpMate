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
    monkeypatch.setattr(windows_integration.sys, "version_info", (3, 14, 0))
    windows_integration.get_native_admin_drag_drop_policy.cache_clear()

    enabled, reason = windows_integration.get_native_admin_drag_drop_policy()

    assert enabled is False
    assert "Python 3.14+" in reason


def test_native_admin_drag_drop_policy_force_env_overrides_default(monkeypatch) -> None:
    monkeypatch.setenv(windows_integration.NATIVE_DND_FORCE_ENV, "1")
    monkeypatch.setattr(windows_integration.sys, "version_info", (3, 14, 0))
    windows_integration.get_native_admin_drag_drop_policy.cache_clear()

    enabled, reason = windows_integration.get_native_admin_drag_drop_policy()

    assert enabled is True
    assert windows_integration.NATIVE_DND_FORCE_ENV in reason
