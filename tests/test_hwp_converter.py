from __future__ import annotations

from pathlib import Path

from hwpmate.services.hwp_converter import HWPConverter


class FakeHwp:
    def __init__(
        self,
        *,
        open_result: bool | int = True,
        save_results=None,
        write_output: bool = True,
        output_content: bytes = b"x",
        aux_suffix: str | None = None,
    ) -> None:
        self.open_result = open_result
        self.save_results = list(save_results or [True])
        self.write_output = write_output
        self.output_content = output_content
        self.aux_suffix = aux_suffix
        self.save_calls: list[tuple[str, str, str | None]] = []
        self.clear_calls: list[int] = []

    def RegisterModule(self, module_name: str, module_name_alias: str):
        del module_name, module_name_alias

    def SetMessageBoxMode(self, mode: int):
        del mode

    def Open(self, path: str, format_name: str, options: str):
        del path, format_name, options
        return self.open_result

    def SaveAs(self, path: str, format_name: str, options: str | None = None):
        self.save_calls.append((path, format_name, options))
        result = self.save_results.pop(0) if self.save_results else True
        if result is True and self.write_output:
            Path(path).write_bytes(self.output_content)
        if result is True and self.aux_suffix:
            output = Path(path)
            aux_path = output.with_name(f"{output.stem}{self.aux_suffix}{output.suffix}")
            aux_path.write_bytes(self.output_content)
        return result

    def Clear(self, option: int = 0):
        self.clear_calls.append(option)

    def Quit(self):
        return None


def build_converter(fake_hwp: FakeHwp) -> HWPConverter:
    converter = HWPConverter()
    converter.hwp = fake_hwp
    converter.is_initialized = True
    return converter


def test_kill_owned_processes_skips_pid_not_in_live_hwp_snapshot(monkeypatch) -> None:
    import hwpmate.services.hwp_converter as converter_module

    converter = HWPConverter()
    converter.owned_pids = {111, 222}
    monkeypatch.setattr(converter_module, "_snapshot_hwp_pids", lambda: {222})
    calls: list[tuple[list[str], object | None]] = []

    def fake_run(args, **kwargs):
        calls.append((list(args), kwargs.get("creationflags")))

        class Result:
            returncode = 0

        return Result()

    monkeypatch.setattr(converter_module.subprocess, "run", fake_run)

    assert converter.kill_owned_processes() is True
    assert len(calls) == 1
    assert calls[0][0] == ["taskkill", "/PID", "222", "/F"]
    assert calls[0][1] == converter_module._CREATE_NO_WINDOW
    assert converter.owned_pids == set()


def test_try_set_xhw_windows_visible_sets_all_items() -> None:
    class FakeWindow:
        def __init__(self) -> None:
            self.Visible = True

    class FakeXWindows:
        def __init__(self) -> None:
            self.Count = 2
            self.items = [FakeWindow(), FakeWindow()]

        def Item(self, index: int):
            return self.items[index]

    class FakeHwpWithWindows(FakeHwp):
        def __init__(self) -> None:
            super().__init__()
            self.XHwpWindows = FakeXWindows()

    fake = FakeHwpWithWindows()
    converter = build_converter(fake)
    assert converter._try_set_xhw_windows_visible(False) is True
    assert all(w.Visible is False for w in fake.XHwpWindows.items)


def test_suppress_hwp_ui_flash_calls_visible_and_windows_helper(monkeypatch) -> None:
    import hwpmate.windows_integration as wi

    fake = FakeHwp()
    converter = build_converter(fake)
    converter.owned_pids = {42}
    visible_calls: list[bool] = []
    helper_calls: list[set[int] | None] = []

    monkeypatch.setattr(
        converter,
        "_try_set_xhw_windows_visible",
        lambda visible: visible_calls.append(visible) or True,
    )

    def fake_suppress(pids=None):
        helper_calls.append(pids)
        return (1, 0)

    monkeypatch.setattr(wi, "suppress_hwp_ui_flash", fake_suppress)

    converter._suppress_hwp_ui_flash()

    assert visible_calls == [False]
    assert helper_calls == [{42}]


def test_cleanup_does_not_uninitialize_unowned_com_apartment(monkeypatch) -> None:
    import hwpmate.services.hwp_converter as converter_module

    class FakePythoncom:
        def __init__(self) -> None:
            self.uninit_calls = 0

        def CoUninitialize(self) -> None:
            self.uninit_calls += 1

    fake = FakePythoncom()
    monkeypatch.setattr(converter_module, "pythoncom", fake)
    converter = HWPConverter()
    converter.is_initialized = True
    converter.hwp = FakeHwp()
    converter._com_apartment_owned = False

    converter.cleanup()

    assert fake.uninit_calls == 0
    assert converter.is_initialized is False


def test_convert_file_fails_when_open_returns_false(tmp_path: Path) -> None:
    source = tmp_path / "a.hwp"
    source.write_text("x", encoding="utf-8")
    output = tmp_path / "a.pdf"
    fake = FakeHwp(open_result=False)

    success, error = build_converter(fake).convert_file(source, output, "PDF")

    assert success is False
    assert error is not None and "문서 열기 실패" in error
    assert not fake.save_calls
    assert fake.clear_calls == [1]


def test_convert_file_fails_when_open_returns_zero(tmp_path: Path) -> None:
    source = tmp_path / "a.hwp"
    source.write_text("x", encoding="utf-8")
    output = tmp_path / "a.pdf"
    fake = FakeHwp(open_result=0)

    success, error = build_converter(fake).convert_file(source, output, "PDF")

    assert success is False
    assert error is not None and "문서 열기 실패" in error
    assert not fake.save_calls


def test_is_com_failure_result_normalizes_false_and_zero() -> None:
    from hwpmate.services.hwp_converter import is_com_failure_result

    assert is_com_failure_result(False) is True
    assert is_com_failure_result(0) is True
    assert is_com_failure_result(True) is False
    assert is_com_failure_result(1) is False
    assert is_com_failure_result(None) is False


def test_convert_file_falls_back_when_saveas_returns_false(tmp_path: Path) -> None:
    source = tmp_path / "a.hwp"
    source.write_text("x", encoding="utf-8")
    output = tmp_path / "a.pdf"
    fake = FakeHwp(save_results=[False, True])

    success, error = build_converter(fake).convert_file(source, output, "PDF")

    assert success is True
    assert error is None
    assert len(fake.save_calls) == 2
    assert fake.save_calls[1][2] == ""


def test_convert_file_fails_when_output_file_is_missing(tmp_path: Path) -> None:
    source = tmp_path / "a.hwp"
    source.write_text("x", encoding="utf-8")
    output = tmp_path / "a.pdf"
    fake = FakeHwp(write_output=False)

    success, error = build_converter(fake).convert_file(source, output, "PDF")

    assert success is False
    assert error is not None and "생성되지 않았습니다" in error


def test_convert_file_fails_when_existing_output_is_not_updated(tmp_path: Path) -> None:
    source = tmp_path / "a.hwp"
    source.write_text("x", encoding="utf-8")
    output = tmp_path / "a.pdf"
    output.write_bytes(b"old")
    fake = FakeHwp(write_output=False)

    success, error = build_converter(fake).convert_file(source, output, "PDF")

    assert success is False
    assert error is not None and "갱신되지 않았습니다" in error


def test_convert_file_fails_when_output_file_is_empty(tmp_path: Path) -> None:
    source = tmp_path / "a.hwp"
    source.write_text("x", encoding="utf-8")
    output = tmp_path / "a.pdf"
    fake = FakeHwp(output_content=b"")

    success, error = build_converter(fake).convert_file(source, output, "PDF")

    assert success is False
    assert error is not None and "비어 있습니다" in error


def test_convert_file_accepts_auxiliary_image_artifact(tmp_path: Path) -> None:
    source = tmp_path / "a.hwp"
    source.write_text("x", encoding="utf-8")
    output = tmp_path / "a.png"
    fake = FakeHwp(write_output=False, aux_suffix="_001")
    converter = build_converter(fake)

    success, error = converter.convert_file(source, output, "PNG")

    assert success is True
    assert error is None
    assert converter.last_created_files == [tmp_path / "a_001.png"]
    assert converter.last_output_size == 1
    assert converter.last_save_format == "PNG"
