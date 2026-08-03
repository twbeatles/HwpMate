from __future__ import annotations

from pathlib import Path

from hwpmate.services import hwp_security_module as sec


def test_ensure_hwp_security_module_writes_registry_and_installs_dll(
    tmp_path: Path, monkeypatch
) -> None:
    dll_src = tmp_path / "FilePathCheckerModuleExample.dll"
    dll_src.write_bytes(b"MZ-fake-dll-content")

    runtime_dir = tmp_path / "runtime_security"
    monkeypatch.setattr(sec, "_runtime_security_dir", lambda: runtime_dir)
    monkeypatch.setattr(sec, "find_bundled_security_dll", lambda: dll_src)
    # 단위 테스트용 가짜 DLL — 무결성 검증을 우회
    monkeypatch.setattr(sec, "verify_dll_integrity", lambda path, **kwargs: None)

    written: dict[str, str] = {}

    def fake_write(dll_path: Path, alias: str = sec.SECURITY_MODULE_ALIAS) -> str:
        written["path"] = str(dll_path)
        written["alias"] = alias
        return alias

    def fake_read(alias: str = sec.SECURITY_MODULE_ALIAS):
        if written.get("alias") != alias:
            return None
        return written.get("path")

    monkeypatch.setattr(sec, "write_security_module_registry", fake_write)
    monkeypatch.setattr(sec, "read_security_module_registry", fake_read)

    ok, msg, alias = sec.ensure_hwp_security_module()

    assert ok is True
    assert alias == sec.SECURITY_MODULE_ALIAS
    assert (runtime_dir / sec.SECURITY_MODULE_DLL_NAME).is_file()
    assert "DLL=" in msg


def test_ensure_hwp_security_module_fails_without_dll(tmp_path: Path, monkeypatch) -> None:
    monkeypatch.setattr(sec, "_runtime_security_dir", lambda: tmp_path / "empty")
    monkeypatch.setattr(sec, "find_bundled_security_dll", lambda: None)

    ok, msg, alias = sec.ensure_hwp_security_module()

    assert ok is False
    assert alias is None
    assert "찾을 수 없" in msg or "DLL" in msg


def test_verify_dll_integrity_rejects_mismatch(tmp_path: Path) -> None:
    dll = tmp_path / "bad.dll"
    dll.write_bytes(b"not-the-expected-content")

    try:
        sec.verify_dll_integrity(dll)
        raised = False
    except ValueError as exc:
        raised = True
        assert "무결성" in str(exc)

    assert raised is True


def test_verify_dll_integrity_accepts_known_hash(tmp_path: Path, monkeypatch) -> None:
    dll = tmp_path / "ok.dll"
    content = b"known-content-for-hash"
    dll.write_bytes(content)
    expected = sec.sha256_file(dll)
    monkeypatch.setattr(sec, "EXPECTED_DLL_SHA256", expected)

    sec.verify_dll_integrity(dll, expected_sha256=expected)
