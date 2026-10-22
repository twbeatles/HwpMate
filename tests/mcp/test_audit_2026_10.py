"""2026-10-22 감사(PROJECT_AUDIT.md ISSUE-001~005, §5 Gap) 회귀 테스트."""

from __future__ import annotations

import json
import os
import subprocess
import time
from pathlib import Path

import pytest

from hwpmate.mcp import __main__ as mcp_main
from hwpmate.mcp import runner
from hwpmate.mcp.capabilities import gather_capabilities
from hwpmate.mcp.config import McpConfig, McpConfigError, load_config
from hwpmate.mcp.errors import McpError
from hwpmate.mcp.jobs import JobManager, recheck_input, wait_for_state
from hwpmate.mcp.planner_adapter import get_plan, preview_conversion
from hwpmate.mcp.policy import check_input_path, check_output_dir
from hwpmate.mcp.schemas import CONFIRMATION_TOKEN
from hwpmate.services.update_installer import wait_for_process_exit

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


def _submit(manager: JobManager, cfg: McpConfig, inputs: list[str], key: str):
    preview = preview_conversion(cfg, input_paths=inputs, format_type="PDF", output_dir=cfg.output_roots[0])
    return manager.submit(get_plan(preview["plan_id"]), confirmation=CONFIRMATION_TOKEN, idempotency_key=key)


@pytest.fixture(autouse=True)
def _clean_fake_env(monkeypatch: pytest.MonkeyPatch) -> None:
    for key in list(os.environ):
        if key.startswith("HWPMATE_FAKECLI_") or key.startswith("HWPMATE_MCP_"):
            monkeypatch.delenv(key, raising=False)


# --- ISSUE-002: 허용 루트 형식 검증 -------------------------------------------


def test_string_roots_are_wrapped_not_iterated(tmp_path: Path) -> None:
    allowed = tmp_path / "allowed"
    allowed.mkdir()
    outside = tmp_path / "outside" / "secret.hwp"
    outside.parent.mkdir()
    outside.write_bytes(b"x")
    cfg = McpConfig.from_mapping({"input_roots": str(allowed), "output_roots": str(allowed)})
    assert cfg.input_roots == [str(allowed)]
    with pytest.raises(McpError) as excinfo:
        check_input_path(cfg, str(outside))
    assert excinfo.value.code == "INPUT_DENIED"
    with pytest.raises(McpError):
        check_output_dir(cfg, str(tmp_path / "elsewhere"))


def test_string_roots_in_toml_file_do_not_open_drive(tmp_path: Path) -> None:
    allowed = tmp_path / "allowed"
    allowed.mkdir()
    toml_path = tmp_path / "config.toml"
    toml_path.write_text(f'[paths]\ninput_roots = "{allowed.as_posix()}"\n', encoding="utf-8")
    cfg = load_config(toml_path)
    with pytest.raises(McpError):
        check_input_path(cfg, str(tmp_path / "x.hwp"))


def test_drive_root_is_rejected_at_load(tmp_path: Path) -> None:
    drive_root = Path(tmp_path.anchor).as_posix()
    with pytest.raises(McpConfigError):
        McpConfig.from_mapping({"input_roots": [drive_root]})


def test_direct_config_with_bad_roots_still_fails_closed(tmp_path: Path) -> None:
    target = tmp_path / "a.hwp"
    target.write_bytes(b"x")
    as_string = McpConfig(input_roots=str(tmp_path))  # type: ignore[arg-type]
    with pytest.raises(McpError):
        check_input_path(as_string, str(target))
    drive_only = McpConfig(input_roots=[tmp_path.anchor])
    with pytest.raises(McpError):
        check_input_path(drive_only, str(target))


@pytest.mark.parametrize(
    "mapping",
    [
        {"output_roots": 5},
        {"input_roots": ["C:/ok", 3]},
        {"max_files_per_job": "7"},
        {"enabled": "false"},
        {"allow_unc_paths": 1},
        {"cli_executable": ["a.exe"]},
    ],
)
def test_wrong_types_are_rejected(mapping: dict) -> None:
    with pytest.raises(McpConfigError):
        McpConfig.from_mapping(mapping)


# --- ISSUE-005: TOML 문법 오류 ------------------------------------------------


def test_bad_toml_raises_config_error_with_path(tmp_path: Path) -> None:
    bad = tmp_path / "bad.toml"
    bad.write_text("[mcp\nenabled = true\n", encoding="utf-8")
    with pytest.raises(McpConfigError) as excinfo:
        load_config(bad)
    assert str(bad) in str(excinfo.value)


def test_config_file_without_tomllib_is_not_silently_ignored(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    import hwpmate.mcp.config as config_module

    config = tmp_path / "config.toml"
    config.write_text('[paths]\ninput_roots = ["C:/a"]\n', encoding="utf-8")
    monkeypatch.setattr(config_module, "tomllib", None)  # Python 3.10 재현
    with pytest.raises(McpConfigError) as excinfo:
        load_config(config)
    assert "3.11" in str(excinfo.value)
    assert load_config(tmp_path / "missing.toml").input_roots == []


def test_main_exits_cleanly_on_bad_config(tmp_path: Path, capsys: pytest.CaptureFixture[str]) -> None:
    bad = tmp_path / "bad.toml"
    bad.write_text('[paths]\ninput_roots = ["C:\\Work"]\n', encoding="utf-8")  # 잘못된 이스케이프
    assert mcp_main.main(["--config", str(bad)]) == 2
    captured = capsys.readouterr()
    assert captured.out == ""  # stdout 은 JSON-RPC 전용
    assert "설정 오류" in captured.err


# --- §5: 적용되지 않는 설정 키 경고 ------------------------------------------


def test_fixed_policy_keys_produce_warnings(tmp_path: Path) -> None:
    cfg = McpConfig.from_mapping({"allow_overwrite": True, "require_confirmation": False})
    assert any("allow_overwrite" in w for w in cfg.config_warnings)
    assert any("require_confirmation" in w for w in cfg.config_warnings)
    caps = gather_capabilities(cfg)
    assert any("allow_overwrite" in w for w in caps["warnings"])
    assert McpConfig.from_mapping({}).config_warnings == []


def test_require_confirmation_false_is_not_honored(tmp_path: Path) -> None:
    src = _seed(tmp_path, ["a.hwp"])
    cfg = _cfg(tmp_path, require_confirmation=False)
    manager = JobManager(cfg)
    preview = preview_conversion(cfg, input_paths=[str(src)], format_type="PDF", output_dir=cfg.output_roots[0])
    with pytest.raises(McpError) as excinfo:
        manager.submit(get_plan(preview["plan_id"]), confirmation="", idempotency_key="k")
    assert excinfo.value.code == "CONFIRMATION_REQUIRED"


# --- ISSUE-001: 권한 상승 필요 실행 파일 --------------------------------------


def test_elevation_required_oserror_is_reported(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    def _raise(*args: object, **kwargs: object) -> None:
        raise OSError(None, "요청한 작업을 수행하려면 권한 상승이 필요합니다", None, 740)

    monkeypatch.setattr(runner.subprocess, "Popen", _raise)
    with pytest.raises(McpError) as excinfo:
        runner.run_cli(
            McpConfig(),
            argv=["C:/HwpMate.exe", "--input", "a.hwp"],
            job_private_dir=tmp_path / "step",
            report_path=tmp_path / "r.json",
            timeout_seconds=5,
        )
    assert excinfo.value.code == "ELEVATION_REQUIRED"


def test_elevation_required_fails_job_with_specific_code(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    src = _seed(tmp_path, ["a.hwp"])
    cfg = _cfg(tmp_path)
    manager = JobManager(cfg)

    def _raise(*args: object, **kwargs: object) -> None:
        raise OSError(None, "elevation", None, 740)

    monkeypatch.setattr(runner.subprocess, "Popen", _raise)
    record = _submit(manager, cfg, [str(src)], "e1")
    done = wait_for_state(manager, record.job_id, timeout_seconds=30)
    assert done.state == "failed"
    assert done.error["code"] == "ELEVATION_REQUIRED"


def test_capabilities_flag_admin_manifest(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    admin_exe = tmp_path / "admin.exe"
    admin_exe.write_bytes(b"MZ" + b"\0" * 1000 + b'<requestedExecutionLevel level="requireAdministrator" uiAccess="false"/>')
    plain_exe = tmp_path / "plain.exe"
    plain_exe.write_bytes(b"MZ" + b'<requestedExecutionLevel level="asInvoker"/>')
    assert runner.executable_requires_admin(admin_exe) is True
    assert runner.executable_requires_admin(plain_exe) is False

    monkeypatch.setattr("hwpmate.mcp.capabilities.current_process_is_elevated", lambda: False)
    caps = gather_capabilities(McpConfig(cli_executable=str(admin_exe)))
    assert caps["cli_path_valid"] is False
    assert caps["cli_requires_elevation"] is True
    caps_plain = gather_capabilities(McpConfig(cli_executable=str(plain_exe)))
    assert caps_plain["cli_path_valid"] is True
    assert caps_plain["cli_requires_elevation"] is False

    monkeypatch.setattr("hwpmate.mcp.capabilities.current_process_is_elevated", lambda: True)
    assert gather_capabilities(McpConfig(cli_executable=str(admin_exe)))["cli_path_valid"] is True


# --- ISSUE-004: 중간 BUSY ----------------------------------------------------


def test_busy_mid_job_records_remaining_inputs(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv("HWPMATE_FAKECLI_BUSY_NAMES", "b.hwp")
    src = _seed(tmp_path, ["a.hwp", "b.hwp", "c.hwp"])
    cfg = _cfg(tmp_path)
    manager = JobManager(cfg)
    record = _submit(manager, cfg, [str(src)], "busy-mid")
    done = wait_for_state(manager, record.job_id, timeout_seconds=60)
    statuses = {Path(item["input"]).name: item["status"] for item in done.items}
    assert statuses == {"a.hwp": "성공", "b.hwp": "실패", "c.hwp": "취소됨"}
    assert done.state == "partially_failed"
    assert done.error["code"] == "BUSY_EXTERNAL_INSTANCE"
    status = done.to_status_dict()
    total = status["success_count"] + status["failed_count"] + status["skipped_count"] + status["canceled_count"]
    assert total == status["requested"] == 3


# --- ISSUE-003: 제한 시간 초과 ------------------------------------------------


def test_timeout_stops_remaining_inputs(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv("HWPMATE_FAKECLI_SLEEP_NAMES", "a.hwp")
    monkeypatch.setenv("HWPMATE_FAKECLI_SLEEP", "30")
    src = _seed(tmp_path, ["a.hwp", "b.hwp"])
    cfg = _cfg(tmp_path, job_timeout_seconds=2)
    manager = JobManager(cfg)
    record = _submit(manager, cfg, [str(src)], "to1")
    done = wait_for_state(manager, record.job_id, timeout_seconds=60)
    assert done.state == "timeout"
    assert done.error["code"] == "TIMEOUT"
    by_name = {Path(item["input"]).name: item for item in done.items}
    assert by_name["a.hwp"]["status"] == "실패"
    assert by_name["b.hwp"]["status"] == "취소됨"
    assert not (tmp_path / "out" / "b.pdf").exists()


class _EveryPid:
    """`owned & live` 에서 owned 를 그대로 돌려주는 대역 (이미지명 검사 우회)."""

    def __rand__(self, other: set[int]) -> set[int]:
        return set(other)


def test_timeout_cleans_owned_processes_outside_cli_tree(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv("HWPMATE_FAKECLI_MODE", "sleep")
    monkeypatch.setenv("HWPMATE_FAKECLI_SLEEP", "30")
    monkeypatch.setenv("HWPMATE_FAKECLI_SPAWN_CHILD", "1")
    monkeypatch.setattr(runner, "_live_hwp_pids", lambda: _EveryPid())
    src = _seed(tmp_path, ["a.hwp"])
    cfg = _cfg(tmp_path, job_timeout_seconds=4)
    manager = JobManager(cfg)
    record = _submit(manager, cfg, [str(src)], "orphan")
    done = wait_for_state(manager, record.job_id, timeout_seconds=60)
    assert done.state == "timeout"
    pids_file = tmp_path / "jobs" / record.job_id / "step-000.pids.json"
    owned = runner.read_owned_pids(pids_file)
    assert len(owned) == 1
    pid = next(iter(owned))
    assert wait_for_process_exit(pid, timeout=10)
    assert any("한글 프로세스" in warning for warning in done.warnings)


def test_normal_exit_leaves_nothing_to_clean(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv("HWPMATE_FAKECLI_SPAWN_CHILD", "1")
    killed: list[int] = []
    monkeypatch.setattr(runner, "_live_hwp_pids", lambda: _EveryPid())
    monkeypatch.setattr(runner, "_terminate_pid", lambda pid: killed.append(pid) or True)
    src = _seed(tmp_path, ["a.hwp"])
    cfg = _cfg(tmp_path)
    manager = JobManager(cfg)
    record = _submit(manager, cfg, [str(src)], "clean")
    done = wait_for_state(manager, record.job_id, timeout_seconds=60)
    assert done.state == "succeeded"
    assert killed == []


def test_cleanup_only_targets_live_hwp_pids(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    pids_file = tmp_path / "p.json"
    pids_file.write_text(json.dumps({"pids": [111, 222, "x", -1, True]}), encoding="utf-8")
    killed: list[int] = []
    monkeypatch.setattr(runner, "_live_hwp_pids", lambda: {222, 333})
    monkeypatch.setattr(runner, "_terminate_pid", lambda pid: killed.append(pid) or True)
    assert runner.cleanup_owned_hwp_processes(pids_file) == [222]
    assert killed == [222]
    assert runner.cleanup_owned_hwp_processes(tmp_path / "missing.json") == []


def test_timeout_kills_process_tree(tmp_path: Path) -> None:
    """venv 런처·onefile 부트로더처럼 자식이 실제 작업을 하는 경우도 함께 종료된다."""
    script = tmp_path / "parent.py"
    marker = tmp_path / "child.pid"
    script.write_text(
        "import subprocess, sys, time\n"
        f"p = subprocess.Popen([sys.executable, '-c', 'import time; time.sleep(60)'])\n"
        f"open(r'{marker}', 'w').write(str(p.pid))\n"
        "time.sleep(60)\n",
        encoding="utf-8",
    )
    import sys

    run = runner.run_cli(
        McpConfig(),
        argv=[sys.executable, str(script)],
        job_private_dir=tmp_path / "step",
        report_path=tmp_path / "r.json",
        timeout_seconds=3,
    )
    assert run.timed_out is True
    deadline = time.time() + 5
    while not marker.exists() and time.time() < deadline:
        time.sleep(0.1)
    child_pid = int(marker.read_text())
    assert wait_for_process_exit(child_pid, timeout=10)


# --- §5: 실행 직전 입력 재검증 ------------------------------------------------


def test_recheck_input_rejects_missing_and_outside(tmp_path: Path) -> None:
    src = _seed(tmp_path, ["a.hwp"])
    cfg = _cfg(tmp_path)
    assert recheck_input(cfg, str(src / "a.hwp")) == (src / "a.hwp").resolve()
    (src / "a.hwp").unlink()
    with pytest.raises(McpError) as excinfo:
        recheck_input(cfg, str(src / "a.hwp"))
    assert excinfo.value.code == "INPUT_NOT_FOUND"
    outside = tmp_path / "other.hwp"
    outside.write_bytes(b"x")
    with pytest.raises(McpError) as excinfo:
        recheck_input(cfg, str(outside))
    assert excinfo.value.code == "INPUT_DENIED"


def test_queued_job_input_removed_before_execution(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv("HWPMATE_FAKECLI_SLEEP_NAMES", "a.hwp")
    monkeypatch.setenv("HWPMATE_FAKECLI_SLEEP", "3")
    src = _seed(tmp_path, ["a.hwp", "b.hwp"])
    cfg = _cfg(tmp_path)
    manager = JobManager(cfg)
    first = _submit(manager, cfg, [str(src / "a.hwp")], "first")
    second = _submit(manager, cfg, [str(src / "b.hwp")], "second")
    assert manager.get(second.job_id).state == "queued"
    (src / "b.hwp").unlink()
    wait_for_state(manager, first.job_id, timeout_seconds=60)
    done = wait_for_state(manager, second.job_id, timeout_seconds=60)
    assert done.state == "failed"
    assert done.items[0]["detail"].startswith("INPUT_NOT_FOUND")


# --- §5: idempotency 색인 손상 복구 -------------------------------------------


def test_corrupt_idempotency_index_is_rebuilt(tmp_path: Path) -> None:
    src = _seed(tmp_path, ["a.hwp"])
    cfg = _cfg(tmp_path)
    manager = JobManager(cfg)
    record = _submit(manager, cfg, [str(src)], "idem")
    wait_for_state(manager, record.job_id, timeout_seconds=60)
    index = tmp_path / "jobs" / "idempotency.json"
    index.write_text('{"broken": ', encoding="utf-8")

    restarted = JobManager(cfg)
    again = _submit(restarted, cfg, [str(src)], "idem")
    assert again.job_id == record.job_id
    assert json.loads(index.read_text(encoding="utf-8"))  # 유효한 JSON으로 다시 저장됨


def test_concurrent_submits_keep_index_valid(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    import threading

    names = [f"f{i}.hwp" for i in range(6)]
    src = _seed(tmp_path, names)
    cfg = _cfg(tmp_path, max_queued_jobs=10)
    manager = JobManager(cfg)
    plans = [
        get_plan(preview_conversion(cfg, input_paths=[str(src / n)], format_type="PDF", output_dir=cfg.output_roots[0])["plan_id"])
        for n in names
    ]
    errors: list[BaseException] = []

    def _go(i: int) -> None:
        try:
            manager.submit(plans[i], confirmation=CONFIRMATION_TOKEN, idempotency_key=f"key-{i}")
        except BaseException as exc:  # pragma: no cover - 실패 시 진단용
            errors.append(exc)

    threads = [threading.Thread(target=_go, args=(i,)) for i in range(len(names))]
    for thread in threads:
        thread.start()
    for thread in threads:
        thread.join()
    assert not errors
    for job_id in list(manager._records):
        wait_for_state(manager, job_id, timeout_seconds=120)
    index = json.loads((tmp_path / "jobs" / "idempotency.json").read_text(encoding="utf-8"))
    assert len(index) == len(names)


def test_subprocess_module_is_real() -> None:
    # 다른 테스트의 monkeypatch 누수 방지 확인용
    assert runner.subprocess is subprocess
