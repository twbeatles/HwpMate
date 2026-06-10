from __future__ import annotations

from pathlib import Path

from hwpmate.services.artifact_policy import (
    existing_artifact_conflicts,
    iter_candidate_artifact_paths,
    matches_artifact_stem,
)


def test_matches_artifact_stem_requires_delimiter_boundary() -> None:
    assert matches_artifact_stem("doc_001.png", "doc")
    assert matches_artifact_stem("doc.files", "doc")
    assert not matches_artifact_stem("document_001.png", "doc")


def test_existing_artifact_conflicts_counts_auxiliary_files_and_directories(tmp_path: Path) -> None:
    output = tmp_path / "doc.png"
    aux_file = tmp_path / "doc_001.png"
    aux_dir = tmp_path / "doc.files"
    unrelated = tmp_path / "document_001.png"
    aux_file.write_bytes(b"x")
    aux_dir.mkdir()
    unrelated.write_bytes(b"x")

    conflicts = existing_artifact_conflicts(output, "PNG")

    assert aux_file in conflicts
    assert aux_dir in conflicts
    assert unrelated not in conflicts


def test_iter_candidate_artifact_paths_limits_nested_scan(tmp_path: Path) -> None:
    output = tmp_path / "doc.html"
    aux_dir = tmp_path / "doc.files"
    second_aux_dir = tmp_path / "doc-assets"
    aux_dir.mkdir()
    second_aux_dir.mkdir()
    for index in range(5):
        (aux_dir / f"{index}.png").write_bytes(b"x")
        (second_aux_dir / f"{index}.png").write_bytes(b"x")

    candidates = iter_candidate_artifact_paths(output, "HTML", nested_limit=2)

    nested = [path for path in candidates if path.parent in {aux_dir, second_aux_dir}]
    assert len(nested) == 2
    assert output in candidates
