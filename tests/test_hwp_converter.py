from __future__ import annotations

from pathlib import Path

from hwpmate.services.hwp_converter import HWPConverter


class FakeHwp:
    def __init__(
        self,
        *,
        open_result=True,
        save_results=None,
        write_output: bool = True,
        output_content: bytes = b"x",
    ) -> None:
        self.open_result = open_result
        self.save_results = list(save_results or [True])
        self.write_output = write_output
        self.output_content = output_content
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


def test_convert_file_fails_when_open_returns_false(tmp_path: Path) -> None:
    source = tmp_path / "a.hwp"
    source.write_text("x", encoding="utf-8")
    output = tmp_path / "a.pdf"
    fake = FakeHwp(open_result=False)

    success, error = build_converter(fake).convert_file(source, output, "PDF")

    assert success is False
    assert error is not None and "문서 열기 실패" in error
    assert not fake.save_calls


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


def test_convert_file_fails_when_output_file_is_empty(tmp_path: Path) -> None:
    source = tmp_path / "a.hwp"
    source.write_text("x", encoding="utf-8")
    output = tmp_path / "a.pdf"
    fake = FakeHwp(output_content=b"")

    success, error = build_converter(fake).convert_file(source, output, "PDF")

    assert success is False
    assert error is not None and "비어 있습니다" in error
