from __future__ import annotations

from pathlib import Path

import pytest

from hwpmate.models import ConversionTask
from hwpmate.services.task_planner import TaskPlanner


def test_build_tasks_in_file_mode_skips_hwpx_when_target_is_hwpx(tmp_path: Path) -> None:
    hwp = tmp_path / "a.hwp"
    hwpx = tmp_path / "b.hwpx"
    hwp.write_text("x", encoding="utf-8")
    hwpx.write_text("x", encoding="utf-8")

    planner = TaskPlanner()
    tasks = planner.build_tasks(
        is_folder_mode=False,
        format_type="HWPX",
        folder_path="",
        include_sub=True,
        same_location=True,
        output_path="",
        file_paths=[str(hwp), str(hwpx)],
    )

    assert [task.input_file for task in tasks] == [hwp]
    assert tasks[0].output_file == hwp.with_suffix(".hwpx")


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
    tasks = planner.build_tasks(
        is_folder_mode=True,
        format_type="PDF",
        folder_path=str(source),
        include_sub=True,
        same_location=False,
        output_path=str(output),
        file_paths=[],
    )

    assert tasks[0].output_file == output / "nested" / "a.pdf"


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
    planner.resolve_output_conflicts(tasks, overwrite=False)
    assert tasks[0].output_file == tmp_path / "doc (2).pdf"

    import hwpmate.services.task_planner as planner_module

    monkeypatch.setattr(planner_module, "MAX_FILENAME_COUNTER", 0)
    tasks = [ConversionTask(input_file=tmp_path / "c.hwp", output_file=existing)]
    planner.resolve_output_conflicts(tasks, overwrite=False)
    assert tasks[0].output_file.name.startswith("doc_")
