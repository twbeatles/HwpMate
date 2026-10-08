from __future__ import annotations

import json
from pathlib import Path

import pytest

from hwpmate.mcp.errors import McpError
from hwpmate.mcp.report_reader import classify_job, read_cli_report


def _report(tmp_path: Path, *, success: int, failed: int, canceled: int = 0, skipped: int = 0) -> Path:
    path = tmp_path / "result.json"
    tasks = []
    for i in range(success):
        tasks.append({"input_file": f"a{i}.hwp", "output_file": f"a{i}.pdf", "status": "성공", "detail": "", "created_files": [f"a{i}.pdf"]})
    for i in range(failed):
        tasks.append({"input_file": f"b{i}.hwp", "output_file": f"b{i}.pdf", "status": "실패", "detail": "boom", "created_files": []})
    for i in range(canceled):
        tasks.append({"input_file": f"c{i}.hwp", "output_file": f"c{i}.pdf", "status": "취소됨", "detail": "", "created_files": []})
    for i in range(skipped):
        tasks.append({"input_file": f"s{i}.pdf", "output_file": f"s{i}.pdf", "status": "건너뜀", "detail": "", "created_files": []})
    payload = {
        "summary": {
            "format_type": "PDF",
            "total_requested": success + failed + canceled + skipped,
            "success_count": success,
            "failed_count": failed,
            "skipped_count": skipped,
            "canceled_count": canceled,
            "elapsed_seconds": 0.1,
            "progid_used": "x",
            "warnings": ["w1"],
        },
        "tasks": tasks,
    }
    path.write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-8")
    return path


def test_read_valid_report(tmp_path: Path) -> None:
    report = read_cli_report(_report(tmp_path, success=2, failed=1))
    assert report["summary"]["success_count"] == 2
    assert report["summary"]["failed_count"] == 1
    assert len(report["items"]) == 3
    assert report["items"][0]["outputs"] == ["a0.pdf"]
    assert report["warnings"] == ["w1"]


def test_read_missing_report(tmp_path: Path) -> None:
    with pytest.raises(McpError) as excinfo:
        read_cli_report(tmp_path / "nope.json")
    assert excinfo.value.code == "REPORT_MISSING"


def test_read_corrupt_report(tmp_path: Path) -> None:
    bad = tmp_path / "bad.json"
    bad.write_text("{not json", encoding="utf-8")
    with pytest.raises(McpError) as excinfo:
        read_cli_report(bad)
    assert excinfo.value.code == "REPORT_MISSING"


def test_read_report_missing_keys(tmp_path: Path) -> None:
    bad = tmp_path / "bad.json"
    bad.write_text(json.dumps({"summary": {}, "tasks": []}), encoding="utf-8")
    with pytest.raises(McpError) as excinfo:
        read_cli_report(bad)
    assert excinfo.value.code == "REPORT_MISSING"


def test_classify_matrix() -> None:
    def overall(success: int, failed: int, canceled: int = 0, skipped: int = 0) -> dict:
        return {"summary": {"success_count": success, "failed_count": failed, "canceled_count": canceled, "skipped_count": skipped}}

    assert classify_job(0, overall(2, 0)) == ("succeeded", None)
    assert classify_job(0, overall(1, 1))[0] == "partially_failed"
    assert classify_job(1, overall(0, 2))[0] == "failed"
    assert classify_job(0, overall(0, 0, canceled=1))[0] == "canceled"
    # 종료 코드만으로 성공을 선언하지 않는다.
    assert classify_job(0, overall(0, 1))[0] == "failed"
