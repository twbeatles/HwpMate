from __future__ import annotations

import time
from pathlib import Path

import pytest

from hwpmate.mcp.config import McpConfig
from hwpmate.mcp.errors import McpError
from hwpmate.mcp.jobs import TERMINAL_STATES, JobManager, wait_for_state
from hwpmate.mcp.planner_adapter import get_plan, preview_conversion
from hwpmate.mcp.schemas import CONFIRMATION_TOKEN

FAKE = str(Path(__file__).parent / "fake_cli.py")


def _cfg(tmp_path: Path, **overrides: object) -> McpConfig:
    base: dict[str, object] = {
        "input_roots": [str(tmp_path / "in")],
        "output_roots": [str(tmp_path / "out")],
        "cli_executable": FAKE,
        "job_dir": str(tmp_path / "jobs"),
        "job_timeout_seconds": 60,
    }
    base.update(overrides)
    return McpConfig.from_mapping(base)


def _seed(tmp_path: Path, names: list[str]) -> Path:
    src = tmp_path / "in"
    src.mkdir(parents=True, exist_ok=True)
    for name in names:
        (src / name).write_bytes(b"dummy")
    return src


def _submit_ok(manager: JobManager, cfg: McpConfig, src: Path, out: str, key: str):
    preview = preview_conversion(cfg, input_paths=[str(src)], format_type="PDF", output_dir=out)
    record = manager.submit(get_plan(preview["plan_id"]), confirmation=CONFIRMATION_TOKEN, idempotency_key=key)
    return record, preview


def _submit_inputs(manager: JobManager, cfg: McpConfig, inputs: list[str], out: str, key: str):
    preview = preview_conversion(cfg, input_paths=inputs, format_type="PDF", output_dir=out)
    return manager.submit(get_plan(preview["plan_id"]), confirmation=CONFIRMATION_TOKEN, idempotency_key=key)


def test_submit_status_result_roundtrip(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv("HWPMATE_FAKECLI_MODE", "ok")
    src = _seed(tmp_path, ["a.hwp", "b.hwpx"])
    cfg = _cfg(tmp_path)
    manager = JobManager(cfg)
    out = str(tmp_path / "out")
    record, _ = _submit_ok(manager, cfg, src, out, "k1")
    assert record.state in ("queued", "running")
    done = wait_for_state(manager, record.job_id, timeout_seconds=60)
    assert done.state == "succeeded"
    assert done.success_count == 2
    assert (tmp_path / "out" / "a.pdf").is_file()
    status = manager.get(record.job_id).to_status_dict()
    assert status["success_count"] == 2


def test_confirmation_required(tmp_path: Path) -> None:
    src = _seed(tmp_path, ["a.hwp"])
    cfg = _cfg(tmp_path)
    manager = JobManager(cfg)
    preview = preview_conversion(cfg, input_paths=[str(src)], format_type="PDF", output_dir=str(tmp_path / "out"))
    with pytest.raises(McpError) as excinfo:
        manager.submit(get_plan(preview["plan_id"]), confirmation="yes", idempotency_key="k")
    assert excinfo.value.code == "CONFIRMATION_REQUIRED"


def test_idempotent_resubmit_returns_same_job(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv("HWPMATE_FAKECLI_MODE", "ok")
    src = _seed(tmp_path, ["a.hwp"])
    cfg = _cfg(tmp_path)
    manager = JobManager(cfg)
    out = str(tmp_path / "out")
    preview = preview_conversion(cfg, input_paths=[str(src)], format_type="PDF", output_dir=out)
    first = manager.submit(get_plan(preview["plan_id"]), confirmation=CONFIRMATION_TOKEN, idempotency_key="same")
    second = manager.submit(get_plan(preview["plan_id"]), confirmation=CONFIRMATION_TOKEN, idempotency_key="same")
    assert first.job_id == second.job_id
    wait_for_state(manager, first.job_id, timeout_seconds=60)


def test_idempotency_conflict_on_different_plan(tmp_path: Path) -> None:
    src = _seed(tmp_path, ["a.hwp", "b.hwp"])
    cfg = _cfg(tmp_path)
    manager = JobManager(cfg)
    out = str(tmp_path / "out")
    preview_a = preview_conversion(cfg, input_paths=[str(src / "a.hwp")], format_type="PDF", output_dir=out)
    preview_b = preview_conversion(cfg, input_paths=[str(src / "b.hwp")], format_type="PDF", output_dir=out)
    manager.submit(get_plan(preview_a["plan_id"]), confirmation=CONFIRMATION_TOKEN, idempotency_key="same")
    with pytest.raises(McpError) as excinfo:
        manager.submit(get_plan(preview_b["plan_id"]), confirmation=CONFIRMATION_TOKEN, idempotency_key="same")
    assert excinfo.value.code == "IDEMPOTENCY_CONFLICT"


def test_queue_full(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv("HWPMATE_FAKECLI_MODE", "sleep")
    monkeypatch.setenv("HWPMATE_FAKECLI_SLEEP", "20")
    src = _seed(tmp_path, ["a.hwp", "b.hwp"])
    cfg = _cfg(tmp_path, max_queued_jobs=0)
    manager = JobManager(cfg)
    out = str(tmp_path / "out")
    _submit_inputs(manager, cfg, [str(src / "a.hwp")], out, "q1")
    preview = preview_conversion(cfg, input_paths=[str(src / "b.hwp")], format_type="PDF", output_dir=out)
    with pytest.raises(McpError) as excinfo:
        manager.submit(get_plan(preview["plan_id"]), confirmation=CONFIRMATION_TOKEN, idempotency_key="q2")
    assert excinfo.value.code == "JOB_QUEUE_FULL"


def test_cancel_queued_job(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv("HWPMATE_FAKECLI_MODE", "sleep")
    monkeypatch.setenv("HWPMATE_FAKECLI_SLEEP", "8")
    src = _seed(tmp_path, ["a.hwp", "b.hwp"])
    cfg = _cfg(tmp_path)
    manager = JobManager(cfg)
    out = str(tmp_path / "out")
    first = _submit_inputs(manager, cfg, [str(src / "a.hwp")], out, "c1")
    second = _submit_inputs(manager, cfg, [str(src / "b.hwp")], out, "c2")
    canceled = manager.cancel(second.job_id, confirmation=CONFIRMATION_TOKEN)
    assert canceled.state == "canceled"
    done = wait_for_state(manager, first.job_id, timeout_seconds=30)
    assert done.state in TERMINAL_STATES


def test_cancel_running_unsupported(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv("HWPMATE_FAKECLI_MODE", "sleep")
    monkeypatch.setenv("HWPMATE_FAKECLI_SLEEP", "8")
    src = _seed(tmp_path, ["a.hwp"])
    cfg = _cfg(tmp_path)
    manager = JobManager(cfg)
    out = str(tmp_path / "out")
    record = _submit_inputs(manager, cfg, [str(src)], out, "r1")
    deadline = time.time() + 10
    while manager.get(record.job_id).state != "running" and time.time() < deadline:
        time.sleep(0.2)
    assert manager.get(record.job_id).state == "running"
    with pytest.raises(McpError) as excinfo:
        manager.cancel(record.job_id, confirmation=CONFIRMATION_TOKEN)
    assert excinfo.value.code == "CANCELLATION_UNSUPPORTED_WHILE_RUNNING"


def test_busy_external_instance(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv("HWPMATE_FAKECLI_MODE", "busy")
    src = _seed(tmp_path, ["a.hwp"])
    cfg = _cfg(tmp_path)
    manager = JobManager(cfg)
    out = str(tmp_path / "out")
    record = _submit_inputs(manager, cfg, [str(src)], out, "b1")
    done = wait_for_state(manager, record.job_id, timeout_seconds=60)
    assert done.state == "rejected_busy"
    assert done.error["code"] == "BUSY_EXTERNAL_INSTANCE"


def test_partial_failure(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv("HWPMATE_FAKECLI_MODE", "ok")
    monkeypatch.setenv("HWPMATE_FAKECLI_FAIL_NAMES", "b.hwp")
    src = _seed(tmp_path, ["a.hwp", "b.hwp"])
    cfg = _cfg(tmp_path)
    manager = JobManager(cfg)
    out = str(tmp_path / "out")
    record = _submit_inputs(manager, cfg, [str(src)], out, "p1")
    done = wait_for_state(manager, record.job_id, timeout_seconds=60)
    assert done.state == "partially_failed"
    assert done.success_count == 1
    assert done.failed_count == 1


def test_report_missing_is_failure(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv("HWPMATE_FAKECLI_MODE", "noreport")
    src = _seed(tmp_path, ["a.hwp"])
    cfg = _cfg(tmp_path)
    manager = JobManager(cfg)
    out = str(tmp_path / "out")
    record = _submit_inputs(manager, cfg, [str(src)], out, "n1")
    done = wait_for_state(manager, record.job_id, timeout_seconds=60)
    assert done.state == "failed"
    assert "REPORT_MISSING" in done.items[0]["detail"]


def test_hancom_unavailable_mapped(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv("HWPMATE_FAKECLI_MODE", "nocom")
    src = _seed(tmp_path, ["a.hwp"])
    cfg = _cfg(tmp_path)
    manager = JobManager(cfg)
    out = str(tmp_path / "out")
    record = _submit_inputs(manager, cfg, [str(src)], out, "h1")
    done = wait_for_state(manager, record.job_id, timeout_seconds=60)
    assert done.state == "failed"
    assert done.error["code"] == "HANCOM_UNAVAILABLE"


def test_queue_expiry(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv("HWPMATE_FAKECLI_MODE", "sleep")
    monkeypatch.setenv("HWPMATE_FAKECLI_SLEEP", "5")
    src = _seed(tmp_path, ["a.hwp", "b.hwp"])
    cfg = _cfg(tmp_path, queue_wait_timeout_seconds=2)
    manager = JobManager(cfg)
    out = str(tmp_path / "out")
    first = _submit_inputs(manager, cfg, [str(src / "a.hwp")], out, "e1")
    second = _submit_inputs(manager, cfg, [str(src / "b.hwp")], out, "e2")
    expired = wait_for_state(manager, second.job_id, timeout_seconds=30)
    assert expired.state == "expired"
    assert expired.error["code"] == "TIMEOUT"
    done = wait_for_state(manager, first.job_id, timeout_seconds=30)
    assert done.state in TERMINAL_STATES


def test_retention_cleanup(tmp_path: Path) -> None:
    from hwpmate.mcp.jobs import JobRecord

    cfg = _cfg(tmp_path, job_retention_days=30)
    manager = JobManager(cfg)
    old = JobRecord(
        job_id="oldjob1234567890",
        format="PDF",
        output_dir=str(tmp_path / "out"),
        state="succeeded",
        created_at="2020-01-01T00:00:00+00:00",
        finished_at="2020-01-02T00:00:00+00:00",
    )
    manager._records[old.job_id] = old
    manager._persist(old)
    stale_tmp = tmp_path / "jobs" / "leftover.tmp"
    stale_tmp.write_bytes(b"x")
    fresh = JobRecord(
        job_id="freshjob123456789",
        format="PDF",
        output_dir=str(tmp_path / "out"),
        state="succeeded",
        created_at="2030-01-01T00:00:00+00:00",
        finished_at="2030-01-01T00:01:00+00:00",
    )
    manager._records[fresh.job_id] = fresh
    manager._persist(fresh)
    reloaded = JobManager(cfg)
    with pytest.raises(McpError) as excinfo:
        reloaded.get(old.job_id)
    assert excinfo.value.code == "JOB_NOT_FOUND"
    assert reloaded.get(fresh.job_id).state == "succeeded"
    assert not stale_tmp.exists()


def test_output_recheck_denied(tmp_path: Path) -> None:
    from hwpmate.mcp.jobs import JobRecord, recheck_output

    cfg = _cfg(tmp_path)
    ok_record = JobRecord(job_id="x", format="PDF", output_dir=str(tmp_path / "out"))
    assert recheck_output(cfg, ok_record).name == "out"
    bad_record = JobRecord(job_id="x", format="PDF", output_dir=str(tmp_path / "evil"))
    with pytest.raises(McpError) as excinfo:
        recheck_output(cfg, bad_record)
    assert excinfo.value.code == "OUTPUT_DENIED"


def test_platform_guard(monkeypatch: pytest.MonkeyPatch) -> None:
    import os as _os

    from hwpmate.mcp.jobs import check_platform_supported

    if _os.name == "nt":
        check_platform_supported()
    monkeypatch.setattr(_os, "name", "posix")
    with pytest.raises(McpError) as excinfo:
        check_platform_supported()
    assert excinfo.value.code == "PLATFORM_UNSUPPORTED"


def test_restart_recovery_marks_interrupted(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv("HWPMATE_FAKECLI_MODE", "sleep")
    monkeypatch.setenv("HWPMATE_FAKECLI_SLEEP", "20")
    src = _seed(tmp_path, ["a.hwp"])
    cfg = _cfg(tmp_path)
    manager = JobManager(cfg)
    out = str(tmp_path / "out")
    record = _submit_inputs(manager, cfg, [str(src)], out, "z1")
    deadline = time.time() + 10
    while manager.get(record.job_id).state not in ("running", "queued") and time.time() < deadline:
        time.sleep(0.1)
    recovered = JobManager(cfg).get(record.job_id)
    assert recovered.state == "interrupted"
