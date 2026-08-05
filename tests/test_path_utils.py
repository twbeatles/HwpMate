from __future__ import annotations

from pathlib import Path

from hwpmate.path_utils import (
    canonicalize_path,
    check_write_permission,
    com_path_candidates,
    is_path_length_blocking,
    is_path_length_risky,
    is_valid_path_name,
    iter_supported_files,
    make_path_key,
    path_char_length,
    to_extended_win_path,
)


def test_canonicalize_and_make_path_key_normalize_windows_paths() -> None:
    raw = r".\docs\..\docs\sample.hwp"

    canonical = canonicalize_path(raw)
    key = make_path_key(raw)

    assert canonical.endswith(str(Path("docs") / "sample.hwp"))
    assert key == make_path_key(canonical.upper())


def test_iter_supported_files_handles_file_and_folder(tmp_path: Path) -> None:
    root_file = tmp_path / "single.hwpx"
    root_file.write_text("x", encoding="utf-8")
    nested = tmp_path / "nested"
    nested.mkdir()
    nested_file = nested / "child.hwp"
    nested_file.write_text("x", encoding="utf-8")
    (nested / "ignore.txt").write_text("x", encoding="utf-8")

    single = list(iter_supported_files(root_file))
    direct = list(iter_supported_files(tmp_path, include_sub=False))
    recursive = list(iter_supported_files(tmp_path, include_sub=True))

    assert single == [root_file]
    assert root_file in direct
    assert nested_file not in direct
    assert nested_file in recursive


def test_iter_supported_files_honors_cancel_checker(tmp_path: Path) -> None:
    first = tmp_path / "a.hwp"
    second = tmp_path / "b.hwp"
    first.write_text("x", encoding="utf-8")
    second.write_text("x", encoding="utf-8")
    calls = {"count": 0}

    def cancel() -> bool:
        calls["count"] += 1
        return calls["count"] > 1

    files = list(iter_supported_files(tmp_path, include_sub=False, cancel_checker=cancel))

    assert len(files) <= 1


def test_iter_supported_files_excludes_nested_backup_dirs_but_allows_backup_root(tmp_path: Path) -> None:
    source = tmp_path / "source"
    source.mkdir()
    root_file = source / "a.hwp"
    root_file.write_text("x", encoding="utf-8")
    backup = source / "backup"
    backup.mkdir()
    backup_file = backup / "a_backup.hwp"
    backup_file.write_text("x", encoding="utf-8")

    recursive = list(iter_supported_files(source, include_sub=True))
    direct_backup = list(iter_supported_files(backup, include_sub=True))

    assert root_file in recursive
    assert backup_file not in recursive
    assert direct_backup == [backup_file]


def test_check_write_permission_uses_temporary_file(tmp_path: Path) -> None:
    assert check_write_permission(tmp_path) is True
    assert not list(tmp_path.glob(".hwpmate_write_test_*"))


def test_is_valid_path_name_rejects_windows_reserved_and_malformed_names() -> None:
    assert not is_valid_path_name("C:/out/CON")
    assert not is_valid_path_name("C:/out/report. ")
    assert not is_valid_path_name("C:/out/report.")
    assert not is_valid_path_name("C:/out/bad:name")
    assert not is_valid_path_name("C:/out/bad\x01name")
    assert is_valid_path_name("C:/out/report folder")
    assert is_valid_path_name("//server/share/report")
    assert is_valid_path_name(r"\\?\C:\out\report")
    assert is_valid_path_name(r"\\?\UNC\server\share\report")


def test_path_char_length_and_risky() -> None:
    short = Path("C:/a.hwp")
    assert path_char_length(short) < 50
    assert is_path_length_risky(short, warn_length=240) is False
    long = Path("C:/") / ("x" * 250) / "doc.hwp"
    assert is_path_length_risky(long, warn_length=240) is True
    assert is_path_length_blocking(long, block_length=260) is True


def test_to_extended_win_path_and_candidates() -> None:
    drive = to_extended_win_path(r"C:\docs\a.hwp")
    assert drive.startswith("\\\\?\\")
    assert "C:" in drive or "c:" in drive.lower()
    cands = com_path_candidates(r"C:\docs\a.hwp")
    assert len(cands) >= 1
    assert any(p.startswith("\\\\?\\") for p in cands) or cands[0].endswith("a.hwp")
