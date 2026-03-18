from __future__ import annotations

from pathlib import Path

from hwpmate.path_utils import canonicalize_path, iter_supported_files, make_path_key


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
