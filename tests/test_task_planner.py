from __future__ import annotations

import datetime
from pathlib import Path

import pytest

from hwpmate.models import ConversionTask
from hwpmate.services.task_planner import TaskPlanner, count_protected_source_renames


def test_build_tasks_in_file_mode_skips_same_format_entries(tmp_path: Path) -> None:
    hwp = tmp_path / "a.hwp"
    hwpx = tmp_path / "b.hwpx"
    hwp.write_text("x", encoding="utf-8")
    hwpx.write_text("x", encoding="utf-8")

    planner = TaskPlanner()
    plan = planner.build_tasks(
        is_folder_mode=False,
        format_type="HWPX",
        folder_path="",
        include_sub=True,
        same_location=True,
        output_path="",
        file_paths=[str(hwp), str(hwpx)],
    )

    assert [task.input_file for task in plan.tasks] == [hwp]
    assert plan.tasks[0].output_file == hwp.with_suffix(".hwpx")
    assert [task.input_file for task in plan.skipped_tasks] == [hwpx]
    assert plan.skipped_tasks[0].status == "건너뜀"


def test_build_tasks_in_folder_mode_uses_cached_file_paths(tmp_path: Path) -> None:
    source = tmp_path / "source"
    source.mkdir()
    doc = source / "a.hwp"
    doc.write_text("x", encoding="utf-8")
    # 디스크에 더 있어도 캐시만 사용해야 한다.
    (source / "ignored.hwp").write_text("x", encoding="utf-8")

    planner = TaskPlanner()
    plan = planner.build_tasks(
        is_folder_mode=True,
        format_type="PDF",
        folder_path=str(source),
        include_sub=True,
        same_location=True,
        output_path="",
        file_paths=[],
        folder_file_paths=[str(doc)],
    )

    assert len(plan.tasks) == 1
    assert plan.tasks[0].input_file.name == "a.hwp"


def test_build_tasks_outside_folder_falls_back_to_flat_output(tmp_path: Path) -> None:
    source = tmp_path / "source"
    source.mkdir()
    outside = tmp_path / "outside"
    outside.mkdir()
    doc = outside / "a.hwp"
    doc.write_text("x", encoding="utf-8")
    output = tmp_path / "out"
    output.mkdir()

    planner = TaskPlanner()
    plan = planner.build_tasks(
        is_folder_mode=True,
        format_type="PDF",
        folder_path=str(source),
        include_sub=True,
        same_location=False,
        output_path=str(output),
        file_paths=[],
        folder_file_paths=[str(doc)],
    )

    assert len(plan.tasks) == 1
    assert plan.tasks[0].output_file == output / "a.pdf"
    assert any("폴더 밖" in w for w in plan.warnings)


def test_build_tasks_in_folder_mode_uses_relative_output_paths(tmp_path: Path) -> None:
    source = tmp_path / "source"
    source.mkdir()
    nested = source / "nested"
    nested.mkdir()
    doc = nested / "a.hwp"
    doc.write_text("x", encoding="utf-8")
    output = tmp_path / "out"
    output.mkdir()

    planner = TaskPlanner()
    plan = planner.build_tasks(
        is_folder_mode=True,
        format_type="PDF",
        folder_path=str(source),
        include_sub=True,
        same_location=False,
        output_path=str(output),
        file_paths=[],
    )

    assert plan.tasks[0].output_file == output / "nested" / "a.pdf"


def test_build_tasks_in_folder_mode_rejects_file_path(tmp_path: Path) -> None:
    source_file = tmp_path / "source.hwp"
    source_file.write_text("x", encoding="utf-8")

    planner = TaskPlanner()
    with pytest.raises(ValueError, match="폴더 경로"):
        planner.build_tasks(
            is_folder_mode=True,
            format_type="PDF",
            folder_path=str(source_file),
            include_sub=True,
            same_location=True,
            output_path="",
            file_paths=[],
        )


def test_preview_allowed_extensions_match_runnable_files() -> None:
    planner = TaskPlanner()

    assert set(planner.preview_allowed_extensions("HWPX")) == {".hwp"}
    assert set(planner.preview_allowed_extensions("HWP")) == {".hwpx"}
    assert set(planner.preview_allowed_extensions("PDF")) == {".hwp", ".hwpx"}


def test_build_tasks_can_return_only_skipped_entries(tmp_path: Path) -> None:
    hwpx = tmp_path / "same.hwpx"
    hwpx.write_text("x", encoding="utf-8")

    planner = TaskPlanner()
    plan = planner.build_tasks(
        is_folder_mode=False,
        format_type="HWPX",
        folder_path="",
        include_sub=True,
        same_location=True,
        output_path="",
        file_paths=[str(hwpx)],
    )

    assert plan.tasks == []
    assert len(plan.skipped_tasks) == 1
    assert "동일 형식" in plan.warnings[0]


def test_resolve_output_conflicts_numbers_and_falls_back_to_timestamp(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    planner = TaskPlanner()
    existing = tmp_path / "doc.pdf"
    existing.write_text("x", encoding="utf-8")
    second_conflict = tmp_path / "doc (1).pdf"
    second_conflict.write_text("x", encoding="utf-8")

    tasks = [
        ConversionTask(input_file=tmp_path / "a.hwp", output_file=existing),
        ConversionTask(input_file=tmp_path / "b.hwp", output_file=existing),
    ]
    renamed_count = planner.resolve_output_conflicts(tasks, overwrite=False)
    assert renamed_count == 2
    assert tasks[0].output_file == tmp_path / "doc (2).pdf"

    import hwpmate.services.task_planner as planner_module

    monkeypatch.setattr(planner_module, "MAX_FILENAME_COUNTER", 0)
    tasks = [ConversionTask(input_file=tmp_path / "c.hwp", output_file=existing)]
    planner.resolve_output_conflicts(tasks, overwrite=False)
    assert tasks[0].output_file.name.startswith("doc_")


def test_resolve_output_conflicts_allows_existing_overwrite_but_renames_batch_duplicates(tmp_path: Path) -> None:
    planner = TaskPlanner()
    existing = tmp_path / "doc.pdf"
    existing.write_text("old", encoding="utf-8")
    tasks = [
        ConversionTask(input_file=tmp_path / "a.hwp", output_file=existing),
        ConversionTask(input_file=tmp_path / "b.hwpx", output_file=existing),
    ]

    renamed_count = planner.resolve_output_conflicts(tasks, overwrite=True)

    assert renamed_count == 1
    assert tasks[0].output_file == existing
    assert tasks[1].output_file == tmp_path / "doc (1).pdf"
    assert tasks[1].conflict_original_output_file == existing


def test_allocate_output_path_renames_late_auxiliary_conflict(tmp_path: Path) -> None:
    planner = TaskPlanner()
    task = ConversionTask(input_file=tmp_path / "source.hwp", output_file=tmp_path / "doc.png")
    (tmp_path / "doc_001.png").write_bytes(b"old")

    changed = planner.allocate_output_path(
        task,
        used_path_keys=set(),
        overwrite=False,
        format_type="PNG",
    )

    assert changed is True
    assert task.output_file == tmp_path / "doc (1).png"
    assert task.conflict_original_output_file == tmp_path / "doc.png"


def test_resolve_output_conflicts_renames_existing_auxiliary_artifact(tmp_path: Path) -> None:
    planner = TaskPlanner()
    existing_aux = tmp_path / "doc_001.png"
    existing_aux.write_bytes(b"old")
    task = ConversionTask(input_file=tmp_path / "a.hwp", output_file=tmp_path / "doc.png")

    renamed_count = planner.resolve_output_conflicts([task], overwrite=False, format_type="PNG")

    assert renamed_count == 1
    assert task.output_file == tmp_path / "doc (1).png"
    assert task.conflict_original_output_file == tmp_path / "doc.png"


def test_resolve_output_conflicts_ignores_unrelated_same_prefix_artifact(tmp_path: Path) -> None:
    planner = TaskPlanner()
    (tmp_path / "document_001.png").write_bytes(b"old")
    task = ConversionTask(input_file=tmp_path / "a.hwp", output_file=tmp_path / "doc.png")

    renamed_count = planner.resolve_output_conflicts([task], overwrite=False, format_type="PNG")

    assert renamed_count == 0
    assert task.output_file == tmp_path / "doc.png"


def test_resolve_output_conflicts_timestamp_fallback_avoids_duplicates(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    planner = TaskPlanner()
    existing = tmp_path / "doc.pdf"
    existing.write_text("x", encoding="utf-8")

    import hwpmate.services.task_planner as planner_module

    class FrozenDateTime:
        @classmethod
        def now(cls):
            return datetime.datetime(2026, 4, 27, 12, 0, 0, 123456)

    class FrozenDateModule:
        datetime = FrozenDateTime

    monkeypatch.setattr(planner_module, "MAX_FILENAME_COUNTER", 0)
    monkeypatch.setattr(planner_module, "datetime", FrozenDateModule)

    tasks = [
        ConversionTask(input_file=tmp_path / "a.hwp", output_file=existing),
        ConversionTask(input_file=tmp_path / "b.hwp", output_file=existing),
    ]

    renamed_count = planner.resolve_output_conflicts(tasks, overwrite=False)

    assert renamed_count == 2
    assert tasks[0].output_file != tasks[1].output_file
    assert tasks[0].conflict_original_output_file == existing
    assert tasks[1].conflict_original_output_file == existing


def test_overwrite_never_targets_existing_source_hwp_document(tmp_path: Path) -> None:
    planner = TaskPlanner()
    (tmp_path / "a.hwp").write_bytes(b"ORIGINAL-HWP")
    (tmp_path / "a.hwpx").write_bytes(b"ORIGINAL-HWPX")

    plan = planner.build_tasks(
        is_folder_mode=True,
        format_type="HWP",
        folder_path=str(tmp_path),
        include_sub=False,
        same_location=True,
        output_path="",
        overwrite=True,
        file_paths=[],
    )
    renamed = planner.resolve_output_conflicts(plan.tasks, overwrite=True, format_type="HWP")

    assert [task.input_file.name for task in plan.skipped_tasks] == ["a.hwp"]
    assert [task.output_file.name for task in plan.tasks] == ["a (1).hwp"]
    assert renamed == 1
    assert count_protected_source_renames(plan.tasks) == 1


def test_overwrite_never_targets_existing_source_hwpx_document(tmp_path: Path) -> None:
    planner = TaskPlanner()
    source = tmp_path / "a.hwp"
    source.write_bytes(b"src")
    (tmp_path / "a.hwpx").write_bytes(b"ORIGINAL-HWPX")
    task = ConversionTask(input_file=source, output_file=tmp_path / "a.hwpx")

    planner.resolve_output_conflicts([task], overwrite=True, format_type="HWPX")

    assert task.output_file == tmp_path / "a (1).hwpx"


def test_overwrite_still_replaces_existing_non_source_output(tmp_path: Path) -> None:
    planner = TaskPlanner()
    source = tmp_path / "a.hwp"
    source.write_bytes(b"src")
    (tmp_path / "a.pdf").write_bytes(b"old pdf")
    task = ConversionTask(input_file=source, output_file=tmp_path / "a.pdf")

    renamed = planner.resolve_output_conflicts([task], overwrite=True, format_type="PDF")

    assert renamed == 0
    assert task.output_file == tmp_path / "a.pdf"


def test_image_and_html_same_location_keep_plain_output_name(tmp_path: Path) -> None:
    planner = TaskPlanner()
    (tmp_path / "report.hwp").write_bytes(b"src")
    (tmp_path / "report_final.hwp").write_bytes(b"src")

    for fmt, ext in (("PNG", ".png"), ("JPG", ".jpg"), ("HTML", ".html")):
        plan = planner.build_tasks(
            is_folder_mode=True,
            format_type=fmt,
            folder_path=str(tmp_path),
            include_sub=False,
            same_location=True,
            output_path="",
            file_paths=[],
        )
        renamed = planner.resolve_output_conflicts(plan.tasks, overwrite=False, format_type=fmt)
        assert renamed == 0, fmt
        assert sorted(task.output_file.name for task in plan.tasks) == [
            f"report{ext}",
            f"report_final{ext}",
        ]


def test_existing_hwp_page_images_trigger_rename(tmp_path: Path) -> None:
    planner = TaskPlanner()
    source = tmp_path / "report.hwp"
    source.write_bytes(b"src")
    (tmp_path / "report001.png").write_bytes(b"old page")
    task = ConversionTask(input_file=source, output_file=tmp_path / "report.png")

    renamed = planner.resolve_output_conflicts([task], overwrite=False, format_type="PNG")

    assert renamed == 1
    assert task.output_file == tmp_path / "report (1).png"
