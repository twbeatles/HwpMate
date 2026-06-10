from __future__ import annotations

import os
from pathlib import Path

AUXILIARY_ARTIFACT_FORMATS = frozenset({"HTML", "PNG", "JPG", "BMP", "GIF"})
AUXILIARY_NAME_DELIMITERS = frozenset({"_", "-", " ", ".", "("})
MAX_AUXILIARY_SCAN_FILES = 500


def uses_auxiliary_artifacts(format_type: str) -> bool:
    return format_type.upper() in AUXILIARY_ARTIFACT_FORMATS


def matches_artifact_stem(name: str, stem: str) -> bool:
    """Return True when a file/directory name belongs to the output stem."""
    name_key = name.lower()
    stem_key = stem.lower()
    if not stem_key:
        return False
    if name_key == stem_key:
        return True
    if not name_key.startswith(stem_key):
        return False
    if len(name_key) == len(stem_key):
        return True
    return name_key[len(stem_key)] in AUXILIARY_NAME_DELIMITERS


def artifact_key(path: Path) -> str:
    return os.path.normcase(str(path.resolve() if path.exists() else path.absolute()))


def iter_candidate_artifact_paths(
    output_file: Path,
    format_type: str,
    *,
    include_nested: bool = True,
    nested_limit: int = MAX_AUXILIARY_SCAN_FILES,
) -> list[Path]:
    candidates: dict[str, Path] = {artifact_key(output_file): output_file}
    if not uses_auxiliary_artifacts(format_type):
        return list(candidates.values())

    parent = output_file.parent
    if not parent.exists():
        return list(candidates.values())

    stem = output_file.stem
    nested_count = 0
    try:
        for child in parent.iterdir():
            if not matches_artifact_stem(child.name, stem):
                continue
            if child.is_file():
                candidates[artifact_key(child)] = child
                continue
            if child.is_dir() and include_nested:
                if nested_count >= nested_limit:
                    continue
                try:
                    for nested in child.rglob("*"):
                        if not nested.is_file():
                            continue
                        if nested_count >= nested_limit:
                            break
                        candidates[artifact_key(nested)] = nested
                        nested_count += 1
                except OSError:
                    continue
    except OSError:
        return list(candidates.values())

    return list(candidates.values())


def existing_artifact_conflicts(output_file: Path, format_type: str) -> list[Path]:
    conflicts: dict[str, Path] = {}
    if output_file.exists():
        conflicts[artifact_key(output_file)] = output_file

    if not uses_auxiliary_artifacts(format_type):
        return list(conflicts.values())

    parent = output_file.parent
    if not parent.exists():
        return list(conflicts.values())

    try:
        for child in parent.iterdir():
            if child == output_file:
                continue
            if matches_artifact_stem(child.name, output_file.stem):
                conflicts[artifact_key(child)] = child
    except OSError:
        return list(conflicts.values())

    return sorted(conflicts.values(), key=lambda path: str(path).lower())
