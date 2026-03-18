from __future__ import annotations

from hwpmate.services.file_selection_store import FileSelectionStore


def test_add_paths_preserves_order_and_deduplicates_case_insensitively() -> None:
    store = FileSelectionStore()

    added = store.add_paths([r"C:\Docs\a.hwp", r"c:\docs\A.hwp", r"C:\Docs\b.hwpx"])

    assert added == [r"C:\Docs\a.hwp", r"C:\Docs\b.hwpx"]
    assert store.paths == [r"C:\Docs\a.hwp", r"C:\Docs\b.hwpx"]
    assert store.count == 2


def test_remove_rows_and_clear() -> None:
    store = FileSelectionStore()
    store.add_paths([r"C:\Docs\a.hwp", r"C:\Docs\b.hwpx", r"C:\Docs\c.hwp"])

    removed = store.remove_rows([1])

    assert removed == [r"C:\Docs\b.hwpx"]
    assert store.paths == [r"C:\Docs\a.hwp", r"C:\Docs\c.hwp"]

    store.clear()
    assert store.paths == []
    assert store.count == 0
