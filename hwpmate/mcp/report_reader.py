"""CLI --report JSON 파싱 (ConversionSummary.to_json_dict 형태).

성공 판정: 종료 코드 0 AND 유효한 report AND 실패·취소 0건.
report 누락·손상은 성공으로 보정하지 않는다.
"""

from __future__ import annotations

import json
from pathlib import Path

from .errors import McpError

_REQUIRED_SUMMARY_KEYS = {
    "format_type",
    "total_requested",
    "success_count",
    "failed_count",
    "skipped_count",
    "canceled_count",
}


def read_cli_report(report_path: Path) -> dict:
    try:
        raw = report_path.read_text(encoding="utf-8")
    except OSError as exc:
        raise McpError("REPORT_MISSING", f"CLI 결과 보고서를 읽을 수 없습니다: {report_path} ({exc})") from exc
    try:
        payload = json.loads(raw)
    except (ValueError, UnicodeError) as exc:
        raise McpError("REPORT_MISSING", f"CLI 결과 보고서가 손상되었습니다: {report_path} ({exc})") from exc
    if not isinstance(payload, dict):
        raise McpError("REPORT_MISSING", f"CLI 결과 보고서 형식이 올바르지 않습니다: {report_path}")
    summary = payload.get("summary")
    tasks = payload.get("tasks")
    if not isinstance(summary, dict) or not isinstance(tasks, list):
        raise McpError("REPORT_MISSING", f"CLI 결과 보고서에 summary/tasks가 없습니다: {report_path}")
    missing = _REQUIRED_SUMMARY_KEYS - set(summary.keys())
    if missing:
        raise McpError(
            "REPORT_MISSING",
            f"CLI 결과 보고서에 필수 키가 없습니다({', '.join(sorted(missing))}): {report_path}",
        )
    items = []
    for task in tasks:
        if not isinstance(task, dict):
            continue
        created = task.get("created_files", [])
        outputs = list(created) if isinstance(created, list) else [str(created)]
        if not outputs:
            outputs = [str(task.get("output_file", ""))]
        items.append(
            {
                "input": str(task.get("input_file", "")),
                "outputs": [str(path) for path in outputs],
                "status": str(task.get("status", "")),
                "detail": str(task.get("detail", "")),
                "retry_count": task.get("retry_count", 0),
            }
        )
    return {
        "summary": {key: summary.get(key) for key in sorted(_REQUIRED_SUMMARY_KEYS)},
        "warnings": list(summary.get("warnings", []) or []),
        "items": items,
    }


def classify_job(returncode: int, report: dict) -> tuple[str, str | None]:
    """(상태, 에러코드) 반환. report 누락은 호출 전 처리된다."""
    summary = report["summary"]
    failed = int(summary.get("failed_count", 0) or 0)
    canceled = int(summary.get("canceled_count", 0) or 0)
    success = int(summary.get("success_count", 0) or 0)
    if returncode == 0 and failed == 0 and canceled == 0:
        return "succeeded", None
    if canceled > 0 and success == 0 and failed == 0:
        return "canceled", None
    if success > 0 and (failed > 0 or canceled > 0):
        return "partially_failed", None
    return "failed", None
