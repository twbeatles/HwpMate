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


def test_matches_artifact_stem_accepts_hwp_image_page_numbers_only_when_allowed() -> None:
    # 한글 2022 실측: 이미지 SaveAs 는 sample001.png, sample002.png 로 저장
    assert matches_artifact_stem("sample001.png", "sample", allow_page_number=True)
    assert not matches_artifact_stem("sample001.png", "sample")
    assert not matches_artifact_stem("sample1.png", "sample", allow_page_number=True)
    assert not matches_artifact_stem("samples.png", "sample", allow_page_number=True)


def test_existing_artifact_conflicts_ignores_source_documents_and_backup_dir(tmp_path: Path) -> None:
    (tmp_path / "report.hwp").write_bytes(b"src")
    (tmp_path / "report_final.hwpx").write_bytes(b"src")
    (tmp_path / "backup").mkdir()

    assert existing_artifact_conflicts(tmp_path / "report.png", "PNG") == []
    assert existing_artifact_conflicts(tmp_path / "report.html", "HTML") == []


def test_image_candidates_require_same_extension_and_include_page_numbers(tmp_path: Path) -> None:
    output = tmp_path / "sample.png"
    page1 = tmp_path / "sample001.png"
    page2 = tmp_path / "sample002.png"
    other_ext = tmp_path / "sample_notes.txt"
    for path in (page1, page2, other_ext):
        path.write_bytes(b"x")

    candidates = iter_candidate_artifact_paths(output, "PNG")

    assert page1 in candidates
    assert page2 in candidates
    assert other_ext not in candidates


def test_image_candidates_ignore_sibling_document_pages(tmp_path: Path) -> None:
    for name in ("report (1)001.png", "report1001.png", "report-2001.png", "report_final001.png"):
        (tmp_path / name).write_bytes(b"x")

    assert existing_artifact_conflicts(tmp_path / "report.png", "PNG") == []


def test_html_embedded_pic_images_are_tracked_but_not_conflicts(tmp_path: Path) -> None:
    output = tmp_path / "doc.html"
    pic = tmp_path / "PIC388B.png"
    pic.write_bytes(b"x")

    assert pic in iter_candidate_artifact_paths(output, "HTML")
    assert existing_artifact_conflicts(output, "HTML") == []
