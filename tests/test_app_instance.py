from __future__ import annotations

from pathlib import Path

from hwpmate.app_instance import SingleInstanceLock


def test_single_instance_lock_blocks_second_holder(tmp_path: Path) -> None:
    lock_path = tmp_path / "HwpMate.lock"
    first = SingleInstanceLock(lock_path)
    second = SingleInstanceLock(lock_path)

    try:
        assert first.try_lock() is True
        assert second.try_lock() is False
    finally:
        first.release()
        second.release()

    assert second.try_lock() is True
    second.release()
