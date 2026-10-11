"""P1-B — MCP 검증기가 exit 0 + 성공 report 를 맹신하지 않는다는 live 증명.

가짜 CLI(ok 모드)는 DOCX 작업에도 "%PDF-1.4 fake" 바이트를 .docx 로 쓴다.
runner/report_reader 가 성공으로 기록해도 validator 는 ZIP 서명 불일치를
실패로 구분해야 한다. 그 뒤 실제 ZIP 바이트로 고치면 통과한다.
"""

from __future__ import annotations

from pathlib import Path

import pytest

from hwpmate.mcp.config import McpConfig
from hwpmate.mcp.jobs import JobManager, wait_for_state
from hwpmate.mcp.planner_adapter import get_plan, preview_conversion
from hwpmate.mcp.schemas import CONFIRMATION_TOKEN
from hwpmate.mcp.validator import validate_job_outputs

FAKE = str(Path(__file__).parent / "fake_cli.py")


def _cfg(tmp_path: Path) -> McpConfig:
    return McpConfig.from_mapping(
        {
            "input_roots": [str(tmp_path / "in")],
            "output_roots": [str(tmp_path / "out")],
            "cli_executable": FAKE,
            "job_dir": str(tmp_path / "jobs"),
            "job_timeout_seconds": 60,
        }
    )


def _run_docx_job(tmp_path: Path, monkeypatch: pytest.MonkeyPatch):
    monkeypatch.setenv("HWPMATE_FAKECLI_MODE", "ok")
    src = tmp_path / "in"
    src.mkdir(parents=True, exist_ok=True)
    (src / "a.hwp").write_bytes(b"x")
    cfg = _cfg(tmp_path)
    manager = JobManager(cfg)
    out = str(tmp_path / "out")
    preview = preview_conversion(cfg, input_paths=[str(src)], format_type="DOCX", output_dir=out)
    record = manager.submit(get_plan(preview["plan_id"]), confirmation=CONFIRMATION_TOKEN, idempotency_key="d1")
    done = wait_for_state(manager, record.job_id, timeout_seconds=60)
    assert done.state == "succeeded"
    return manager, done


def test_validate_distrusts_success_report_with_wrong_magic(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    """성공 report + exit 0 이어도 내용이 PDF 면 DOCX 검증은 실패한다."""
    manager, done = _run_docx_job(tmp_path, monkeypatch)
    outputs = done.items[0]["outputs"]
    assert outputs[0].endswith(".docx")
    assert Path(outputs[0]).read_bytes().startswith(b"%PDF-1.4")
    result = validate_job_outputs(manager, done.job_id)
    assert result["all_verified"] is False
    assert any("ZIP" in reason for reason in result["invalid"][0]["reasons"])


def test_validate_passes_once_content_matches_report(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    """내용이 report 와 일치하면 같은 job 이 검증 통과한다 (내용 기준 증명)."""
    manager, done = _run_docx_job(tmp_path, monkeypatch)
    for output in done.items[0]["outputs"]:
        Path(output).write_bytes(b"PK\x03\x04" + b"\x00" * 64)
    result = validate_job_outputs(manager, done.job_id)
    assert result["all_verified"] is True
    assert result["verified"] == 1
