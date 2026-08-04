from __future__ import annotations

from pathlib import Path

from hwpmate.services.hwp_print_settings import (
    EXPORT_METHOD_PRINT_TO_PDF_EX,
    EXPORT_METHOD_RUN_TO_PDF,
    PRINT_METHOD_NORMAL,
    apply_default_print_settings,
    is_valid_pdf_file,
    normalize_pdf_export_mode,
    remove_incomplete_output,
    resolve_pdf_printer_candidates,
    try_export_pdf_via_print_to_pdf_ex,
    uses_print_settings_control,
)


def test_uses_print_settings_control_for_pdf_and_images() -> None:
    assert uses_print_settings_control("PDF") is True
    assert uses_print_settings_control("png") is True
    assert uses_print_settings_control("DOCX") is False
    assert uses_print_settings_control("HWP") is False


def test_normalize_pdf_export_mode() -> None:
    assert normalize_pdf_export_mode("saveas_first") == "saveas_first"
    assert normalize_pdf_export_mode("print_to_pdf_ex_first") == "print_to_pdf_ex_first"
    assert normalize_pdf_export_mode("bogus") == "saveas_first"
    assert normalize_pdf_export_mode(None) == "saveas_first"


class _FakeParamSet:
    def __init__(self) -> None:
        self.items: dict[str, object] = {}
        self.HSet = self

    def SetItem(self, key: str, value: object) -> None:
        self.items[key] = value


class _FakeAction:
    def __init__(self, pset: _FakeParamSet) -> None:
        self._pset = pset
        self.get_default_calls: list[str] = []
        self.execute_calls: list[_FakeParamSet] = []

    def CreateSet(self) -> _FakeParamSet:
        return self._pset

    def GetDefault(self, pset: _FakeParamSet) -> None:
        self.get_default_calls.append("Print")
        del pset

    def Execute(self, pset: _FakeParamSet) -> bool:
        self.execute_calls.append(pset)
        return True


class _FakeHAction:
    def __init__(self, pset: _FakeParamSet) -> None:
        self._pset = pset
        self.get_default_calls: list[str] = []
        self.execute_calls: list[tuple[str, _FakeParamSet]] = []

    def GetDefault(self, action_id: str, hset: _FakeParamSet) -> None:
        self.get_default_calls.append(action_id)
        del hset

    def Execute(self, action_id: str, hset: _FakeParamSet) -> bool:
        self.execute_calls.append((action_id, hset))
        return True


class _FakeXPrint:
    def __init__(self) -> None:
        self.PrintMethod = 4
        self.NumCopy = 2
        self.ReverseOrder = 1
        self.filename = ""
        self.PrinterName = ""
        self.run_to_pdf_calls = 0

    def RunToPDF(self) -> None:
        self.run_to_pdf_calls += 1


class _FakeDoc:
    def __init__(self) -> None:
        self.XHwpPrint = _FakeXPrint()


class _FakeDocs:
    def __init__(self, doc: _FakeDoc) -> None:
        self._doc = doc

    def Item(self, index: int) -> _FakeDoc:
        del index
        return self._doc


class FakePrintHwp:
    """인쇄 리셋·PrintToPDFEx 를 흉내 내는 COM 스텁."""

    def __init__(self) -> None:
        self._pset = _FakeParamSet()
        self.HParameterSet = type("HP", (), {"HPrint": self._pset})()
        self.HAction = _FakeHAction(self._pset)
        self._doc = _FakeDoc()
        self.XHwpDocuments = _FakeDocs(self._doc)
        self._print_action = _FakeAction(self._pset)
        self.create_action_calls: list[str] = []

    def CreateAction(self, name: str) -> _FakeAction:
        self.create_action_calls.append(name)
        return self._print_action


def test_apply_default_print_settings_sets_print_method_zero() -> None:
    hwp = FakePrintHwp()
    assert apply_default_print_settings(hwp) is True
    assert hwp._doc.XHwpPrint.PrintMethod == PRINT_METHOD_NORMAL
    assert hwp._doc.XHwpPrint.NumCopy == 1
    assert hwp._doc.XHwpPrint.ReverseOrder == 0
    assert hwp._pset.items.get("PrintMethod") == PRINT_METHOD_NORMAL
    # Execute 하지 않음 (실제 인쇄/파일 생성 방지)
    assert hwp.HAction.execute_calls == []
    assert hwp._print_action.execute_calls == []


def test_apply_default_print_settings_tolerates_minimal_hwp() -> None:
    class Empty:
        pass

    assert apply_default_print_settings(Empty()) is False
    assert apply_default_print_settings(None) is False


def test_try_export_never_executes_create_action_print(tmp_path: Path) -> None:
    """감사 Critical: CreateAction('Print').Execute 는 호출되면 안 된다."""
    hwp = FakePrintHwp()
    out = tmp_path / "out.pdf"

    ok, method = try_export_pdf_via_print_to_pdf_ex(hwp, out)
    assert ok is False
    assert method is None
    assert hwp._print_action.execute_calls == []
    # CreateAction 은 GetDefault 용으로만 쓰일 수 있으나 export 경로에서는 사용 안 함
    # (export 는 HAction.PrintToPDFEx / RunToPDF 만)


def test_try_export_pdf_via_print_to_pdf_ex_writes_file(tmp_path: Path) -> None:
    hwp = FakePrintHwp()
    out = tmp_path / "out.pdf"

    original_execute = hwp.HAction.Execute

    def execute_and_write(action_id: str, hset: _FakeParamSet) -> bool:
        result = original_execute(action_id, hset)
        path = hset.items.get("FileName") or hset.items.get("filename")
        if isinstance(path, str):
            Path(path).write_bytes(b"%PDF-1.4 fake content")
        return result

    hwp.HAction.Execute = execute_and_write  # type: ignore[method-assign]

    ok, method = try_export_pdf_via_print_to_pdf_ex(hwp, out)
    assert ok is True
    assert method == EXPORT_METHOD_PRINT_TO_PDF_EX
    assert out.exists() and is_valid_pdf_file(out)
    assert hwp.HAction.execute_calls
    assert hwp.HAction.execute_calls[0][0] == "PrintToPDFEx"
    assert hwp._pset.items.get("PrintMethod") == PRINT_METHOD_NORMAL
    assert hwp._print_action.execute_calls == []


def test_try_export_rejects_non_pdf_magic(tmp_path: Path) -> None:
    hwp = FakePrintHwp()
    out = tmp_path / "out.pdf"

    def execute_and_write(action_id: str, hset: _FakeParamSet) -> bool:
        del action_id
        path = hset.items.get("FileName") or hset.items.get("filename")
        if isinstance(path, str):
            Path(path).write_bytes(b"NOT-A-PDF-FILE!!!!")
        return True

    hwp.HAction.Execute = execute_and_write  # type: ignore[method-assign]
    ok, method = try_export_pdf_via_print_to_pdf_ex(hwp, out)
    assert ok is False
    assert method is None
    assert not out.exists()  # 불완전 파일 정리


def test_try_export_respects_cancel(tmp_path: Path) -> None:
    hwp = FakePrintHwp()
    out = tmp_path / "out.pdf"
    ok, method = try_export_pdf_via_print_to_pdf_ex(
        hwp, out, cancel_check=lambda: True
    )
    assert ok is False
    assert method is None
    assert hwp.HAction.execute_calls == []


def test_try_export_run_to_pdf_path(tmp_path: Path) -> None:
    hwp = FakePrintHwp()
    out = tmp_path / "out.pdf"

    # HAction 실패, RunToPDF 성공
    def execute_fail(action_id: str, hset: _FakeParamSet) -> bool:
        del action_id, hset
        raise RuntimeError("no PrintToPDFEx")

    hwp.HAction.Execute = execute_fail  # type: ignore[method-assign]

    def run_to_pdf() -> None:
        out.write_bytes(b"%PDF-run")

    hwp._doc.XHwpPrint.RunToPDF = run_to_pdf  # type: ignore[method-assign]

    ok, method = try_export_pdf_via_print_to_pdf_ex(hwp, out)
    assert ok is True
    assert method == EXPORT_METHOD_RUN_TO_PDF


def test_remove_incomplete_output_keeps_unchanged_existing(tmp_path: Path) -> None:
    path = tmp_path / "keep.pdf"
    path.write_bytes(b"%PDF-old")
    st = path.stat()
    remove_incomplete_output(path, before_mtime_ns=st.st_mtime_ns, before_size=st.st_size)
    assert path.exists()


def test_is_valid_pdf_file(tmp_path: Path) -> None:
    good = tmp_path / "g.pdf"
    good.write_bytes(b"%PDF-1.4 x")
    bad = tmp_path / "b.pdf"
    bad.write_bytes(b"xxxx")
    assert is_valid_pdf_file(good) is True
    assert is_valid_pdf_file(bad) is False


def test_resolve_pdf_printer_candidates_fallback() -> None:
    names = resolve_pdf_printer_candidates()
    assert isinstance(names, list)
    assert len(names) >= 1
