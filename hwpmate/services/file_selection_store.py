from __future__ import annotations

from typing import Iterable

from ..path_utils import canonicalize_path, make_path_key


class FileSelectionStore:
    def __init__(self) -> None:
        self._paths: list[str] = []
        self._path_keys: set[str] = set()

    @property
    def paths(self) -> list[str]:
        return self._paths

    @property
    def path_keys(self) -> set[str]:
        return self._path_keys

    @property
    def count(self) -> int:
        return len(self._paths)

    def add_paths(self, paths: Iterable[str]) -> list[str]:
        added: list[str] = []
        for raw_path in paths:
            normalized = canonicalize_path(raw_path)
            key = make_path_key(normalized)
            if key in self._path_keys:
                continue
            self._path_keys.add(key)
            self._paths.append(normalized)
            added.append(normalized)
        return added

    def remove_rows(self, rows: Iterable[int]) -> list[str]:
        removed: list[str] = []
        for row in sorted(set(rows), reverse=True):
            if 0 <= row < len(self._paths):
                path = self._paths.pop(row)
                self._path_keys.discard(make_path_key(path))
                removed.append(path)
        removed.reverse()
        return removed

    def clear(self) -> None:
        self._paths.clear()
        self._path_keys.clear()
