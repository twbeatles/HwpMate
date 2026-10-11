"""P0-D — GUI/CLI/MCP 계획 parity 테스트 (hermetic).

같은 합성 입력 집합을 GUI·CLI 공용 TaskPlanner 와 MCP preview 에 넣고
변환/건너뜀 분류와 출력 파일명 집합이 일치하는지 검증한다.
MCP 보안 계약(overwrite=false 강제, 명시 출력 폴더)은 그대로 둔다.
"""

from __future__ import annotations

from pathlib import Path

from hwpmate.mcp.config import McpConfig
from hwpmate.mcp.errors import McpError
from hwpmate.mcp.planner_adapter import preview_conversion
from hwpmate.services.task_planner import TaskPlanner


def preview_outputs(preview: dict) -> set[str]:
    return {Path(item["output"]).name for item in preview["preview"]}


def _seed(src: Path, names: list[str]) -> None:
    src.mkdir(parents=True, exist_ok=True)
    for name in names:
        (src / name).write_bytes(b"dummy-hwp")


def _cfg(tmp_path: Path) -> McpConfig:
    return McpConfig.from_mapping(
        {
            "input_roots": [str(tmp_path / "in")],
            "output_roots": [str(tmp_path / "out")],
        }
    )


def _planner_names(plan) -> tuple[set[str], set[str], set[str]]:
    planned_inputs = {Path(task.input_file).name for task in plan.tasks}
    planned_outputs = {Path(task.output_file).name for task in plan.tasks}
    skipped_inputs = {Path(task.input_file).name for task in plan.skipped_tasks}
    return planned_inputs, planned_outputs, skipped_inputs


def test_folder_plan_matches_mcp_preview_for_pdf(tmp_path: Path) -> None:
    """PDF 목표: 변환 2건 분류·출력명 일치, 비한글 파일은 양쪽 모두 제외."""
    _seed(tmp_path / "in", ["a.hwp", "b.hwpx", "note.txt", "c.pdf"])
    out = tmp_path / "out"
    out.mkdir()

    plan = TaskPlanner().build_tasks(
        is_folder_mode=True,
        format_type="PDF",
        folder_path=str(tmp_path / "in"),
        include_sub=False,
        same_location=False,
        output_path=str(out),
        overwrite=False,
        file_paths=[],
    )
    preview = preview_conversion(
        _cfg(tmp_path),
        input_paths=[str(tmp_path / "in")],
        format_type="PDF",
        output_dir=str(out),
    )

    planned_inputs, planned_outputs, skipped_inputs = _planner_names(plan)
    assert planned_inputs == {"a.hwp", "b.hwpx"}
    assert skipped_inputs == set()
    assert preview["planned"] == len(plan.tasks) == 2
    assert preview["skipped"] == len(plan.skipped_tasks) == 0
    preview_outputs = {Path(item["output"]).name for item in preview["preview"]}
    assert preview_outputs == planned_outputs == {"a.pdf", "b.pdf"}
    assert {Path(item["input"]).name for item in preview["preview"]} == planned_inputs


def test_same_format_skip_matches_mcp_preview_for_hwp(tmp_path: Path) -> None:
    """HWP 목표: a.hwp 건너뜀·b.hwpx 변환 분류가 양쪽에서 일치."""
    _seed(tmp_path / "in", ["a.hwp", "b.hwpx"])
    out = tmp_path / "out"
    out.mkdir()

    plan = TaskPlanner().build_tasks(
        is_folder_mode=True,
        format_type="HWP",
        folder_path=str(tmp_path / "in"),
        include_sub=False,
        same_location=False,
        output_path=str(out),
        overwrite=False,
        file_paths=[],
    )
    preview = preview_conversion(
        _cfg(tmp_path),
        input_paths=[str(tmp_path / "in")],
        format_type="HWP",
        output_dir=str(out),
    )

    planned_inputs, planned_outputs, skipped_inputs = _planner_names(plan)
    assert planned_inputs == {"b.hwpx"}
    assert skipped_inputs == {"a.hwp"}
    assert preview["planned"] == 1 and preview["requested"] == 2
    assert preview["skipped"] == 1
    assert preview_outputs(preview) == planned_outputs == {"b.hwp"}
    # MCP 는 읽기 전용이다: 출력 폴더 안에 파일을 만들지 않는다.
    assert list(out.iterdir()) == []


def test_mcp_preview_enforces_explicit_output_dir(tmp_path: Path) -> None:
    """MCP 보안 계약: 출력 폴더 미지정은 거부 (fail closed)."""
    _seed(tmp_path / "in", ["a.hwp"])
    try:
        preview_conversion(
            _cfg(tmp_path),
            input_paths=[str(tmp_path / "in")],
            format_type="PDF",
            output_dir="",
        )
    except McpError as exc:
        assert exc.code == "OUTPUT_DENIED"
    else:  # pragma: no cover
        raise AssertionError("OUTPUT_DENIED expected")
