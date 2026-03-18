from __future__ import annotations

import json
from pathlib import Path

from hwpmate.models import ConversionSummary, ConversionTask
from hwpmate.ui.dialogs import write_failed_list, write_results_csv, write_results_json


def build_summary(tmp_path: Path) -> ConversionSummary:
    success = ConversionTask(tmp_path / "a.hwp", tmp_path / "a.pdf", status="성공")
    failed = ConversionTask(tmp_path / "b.hwp", tmp_path / "b.pdf", status="실패", error="save failed")
    skipped = ConversionTask(tmp_path / "c.hwpx", tmp_path / "c.hwpx", status="건너뜀", error="이미 HWPX 형식입니다.")
    canceled = ConversionTask(tmp_path / "d.hwp", tmp_path / "d.pdf", status="취소됨", error="사용자 취소")
    return ConversionSummary(
        format_type="PDF",
        tasks=[success, failed, skipped, canceled],
        warnings=["동일 형식 1개는 자동으로 건너뜁니다."],
        elapsed_seconds=1.25,
        progid_used="Stub.Hwp",
    )


def test_write_failed_list_only_contains_failed_entries(tmp_path: Path) -> None:
    summary = build_summary(tmp_path)
    output = tmp_path / "failed.txt"

    write_failed_list(output, summary.failed_tasks)
    text = output.read_text(encoding="utf-8")

    assert "b.hwp" in text
    assert "save failed" in text
    assert "c.hwpx" not in text


def test_write_results_csv_contains_all_statuses(tmp_path: Path) -> None:
    summary = build_summary(tmp_path)
    output = tmp_path / "results.csv"

    write_results_csv(output, summary)
    text = output.read_text(encoding="utf-8-sig")

    assert "input_file,output_file,status,detail" in text
    assert "성공" in text
    assert "실패" in text
    assert "건너뜀" in text
    assert "취소됨" in text


def test_write_results_json_contains_summary_and_tasks(tmp_path: Path) -> None:
    summary = build_summary(tmp_path)
    output = tmp_path / "results.json"

    write_results_json(output, summary)
    data = json.loads(output.read_text(encoding="utf-8"))

    assert data["summary"]["success_count"] == 1
    assert data["summary"]["failed_count"] == 1
    assert data["summary"]["skipped_count"] == 1
    assert data["summary"]["canceled_count"] == 1
    assert {task["status"] for task in data["tasks"]} == {"성공", "실패", "건너뜀", "취소됨"}
