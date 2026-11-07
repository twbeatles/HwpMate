from __future__ import annotations

from pathlib import Path

from hwpmate.services.artifact_policy import (
    existing_artifact_conflicts,
    is_html_asset_dir_name,
    iter_candidate_artifact_paths,
)


def test_is_html_asset_dir_name_accepts_only_exact_asset_folder_names() -> None:
    assert is_html_asset_dir_name("doc.files", "doc")
    assert is_html_asset_dir_name("DOC_files", "doc")
    assert not is_html_asset_dir_name("doc_2024", "doc")
    assert not is_html_asset_dir_name("document_files", "doc")


def test_existing_artifact_conflicts_counts_page_files_but_not_directories_for_images(tmp_path: Path) -> None:
    output = tmp_path / "doc.png"
    aux_file = tmp_path / "doc_001.png"
    aux_dir = tmp_path / "doc.files"
    unrelated = tmp_path / "document_001.png"
    aux_file.write_bytes(b"x")
    aux_dir.mkdir()
    unrelated.write_bytes(b"x")

    conflicts = existing_artifact_conflicts(output, "PNG")

    assert aux_file in conflicts
    assert aux_dir not in conflicts
    assert unrelated not in conflicts


def test_existing_artifact_conflicts_counts_html_asset_dir(tmp_path: Path) -> None:
    aux_dir = tmp_path / "doc_files"
    aux_dir.mkdir()

    assert existing_artifact_conflicts(tmp_path / "doc.html", "HTML") == [aux_dir]


def test_sibling_folder_does_not_rename_image_output(tmp_path: Path) -> None:
    """감사 ISSUE-003 재현: report_2024/ 폴더 때문에 report (1).png 로 바뀌면 안 된다."""
    (tmp_path / "report.hwp").write_bytes(b"src")
    sibling = tmp_path / "report_2024"
    sibling.mkdir()
    (sibling / "notes.hwp").write_bytes(b"src")
    (sibling / "other.txt").write_bytes(b"x")
    output = tmp_path / "report.png"

    assert existing_artifact_conflicts(output, "PNG") == []
    candidates = iter_candidate_artifact_paths(output, "PNG")
    assert all(path.parent == tmp_path for path in candidates)


def test_sibling_html_file_does_not_rename_html_output(tmp_path: Path) -> None:
    """감사 ISSUE-003 재현: doc_final.html 이 있어도 doc.html 은 그대로여야 한다."""
    (tmp_path / "doc.hwp").write_bytes(b"src")
    (tmp_path / "doc_final.html").write_bytes(b"x")

    assert existing_artifact_conflicts(tmp_path / "doc.html", "HTML") == []


def test_nested_html_asset_scan_excludes_source_documents(tmp_path: Path) -> None:
    aux_dir = tmp_path / "doc_files"
    aux_dir.mkdir()
    (aux_dir / "keep.png").write_bytes(b"x")
    (aux_dir / "inner.hwp").write_bytes(b"src")

    names = {path.name for path in iter_candidate_artifact_paths(tmp_path / "doc.html", "HTML")}

    assert "keep.png" in names
    assert "inner.hwp" not in names


def test_iter_candidate_artifact_paths_limits_nested_scan(tmp_path: Path) -> None:
    output = tmp_path / "doc.html"
    aux_dir = tmp_path / "doc.files"
    second_aux_dir = tmp_path / "doc_files"
    aux_dir.mkdir()
    second_aux_dir.mkdir()
    for index in range(5):
        (aux_dir / f"{index}.png").write_bytes(b"x")
        (second_aux_dir / f"{index}.png").write_bytes(b"x")

    candidates = iter_candidate_artifact_paths(output, "HTML", nested_limit=2)

    nested = [path for path in candidates if path.parent in {aux_dir, second_aux_dir}]
    assert len(nested) == 2
    assert output in candidates


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
