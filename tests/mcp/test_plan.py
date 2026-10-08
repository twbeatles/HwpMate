from __future__ import annotations

from pathlib import Path

import pytest

from hwpmate.mcp import MCP_SCHEMA_VERSION
from hwpmate.mcp.config import McpConfig
from hwpmate.mcp.errors import McpError
from hwpmate.mcp.planner_adapter import get_plan, preview_conversion


def _cfg(tmp_path: Path, **overrides: object) -> McpConfig:
    base: dict[str, object] = {
        "input_roots": [str(tmp_path / "in")],
        "output_roots": [str(tmp_path / "out")],
    }
    base.update(overrides)
    return McpConfig.from_mapping(base)


def _seed(tmp_path: Path, names: list[str]) -> Path:
    src = tmp_path / "in"
    src.mkdir(parents=True, exist_ok=True)
    for name in names:
        target = src / name
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_bytes(b"dummy-hwp")
    return src


def test_preview_counts_and_skips(tmp_path: Path) -> None:
    src = _seed(tmp_path, ["a.hwp", "b.hwpx", "c.pdf", "same.pdf"])
    cfg = _cfg(tmp_path)
    result = preview_conversion(
        cfg, input_paths=[str(src)], format_type="PDF", output_dir=str(tmp_path / "out")
    )
    assert result["schema_version"] == MCP_SCHEMA_VERSION
    assert result["format"] == "PDF"
    assert result["requested"] == 2  # hwp/hwp만 수집
    assert result["planned"] == 2
    assert result["skipped"] == 0
    assert result["requires_confirmation"] is True
    assert result["plan_id"]
    assert len(result["preview"]) == 2


def test_preview_same_format_skipped(tmp_path: Path) -> None:
    src = _seed(tmp_path, ["a.hwp"])
    cfg = _cfg(tmp_path)
    result = preview_conversion(
        cfg, input_paths=[str(src)], format_type="HWP", output_dir=str(tmp_path / "out")
    )
    assert result["requested"] == 1
    assert result["planned"] == 0
    assert result["skipped"] == 1


def test_preview_is_read_only(tmp_path: Path) -> None:
    src = _seed(tmp_path, ["a.hwp"])
    cfg = _cfg(tmp_path)
    out_dir = tmp_path / "out"
    before = list(src.rglob("*"))
    result = preview_conversion(
        cfg, input_paths=[str(src)], format_type="PDF", output_dir=str(out_dir)
    )
    assert result["ok"] is True
    # 출력 폴더를 만들지 않고 입력 트리도 변경하지 않는다.
    assert not out_dir.exists()
    assert list(src.rglob("*")) == before


def test_preview_excludes_backup_and_hidden(tmp_path: Path) -> None:
    src = _seed(tmp_path, ["a.hwp", "backup/old.hwp", ".hidden/s.hwp"])
    (src / "backup").mkdir(parents=True, exist_ok=True)
    (src / "backup" / "old.hwp").write_bytes(b"x")
    cfg = _cfg(tmp_path)
    result = preview_conversion(
        cfg, input_paths=[str(src)], format_type="PDF", output_dir=str(tmp_path / "out")
    )
    assert result["requested"] == 1


def test_preview_conflict_rename(tmp_path: Path) -> None:
    src = _seed(tmp_path, ["a.hwp"])
    out_dir = tmp_path / "out"
    out_dir.mkdir()
    (out_dir / "a.pdf").write_bytes(b"existing")
    cfg = _cfg(tmp_path)
    result = preview_conversion(
        cfg, input_paths=[str(src)], format_type="PDF", output_dir=str(out_dir)
    )
    assert result["conflicts_renamed"] == 1
    assert result["preview"][0]["output"].endswith("a (1).pdf")


def test_preview_quota_enforced(tmp_path: Path) -> None:
    src = _seed(tmp_path, ["a.hwp", "b.hwp"])
    cfg = _cfg(tmp_path, max_files_per_job=1)
    with pytest.raises(McpError) as excinfo:
        preview_conversion(
            cfg, input_paths=[str(src)], format_type="PDF", output_dir=str(tmp_path / "out")
        )
    assert excinfo.value.code == "INPUT_TOO_LARGE"


def test_plan_stale_on_change(tmp_path: Path) -> None:
    src = _seed(tmp_path, ["a.hwp"])
    cfg = _cfg(tmp_path)
    result = preview_conversion(
        cfg, input_paths=[str(src)], format_type="PDF", output_dir=str(tmp_path / "out")
    )
    (src / "a.hwp").write_bytes(b"changed-content-longer")
    with pytest.raises(McpError) as excinfo:
        get_plan(result["plan_id"])
    assert excinfo.value.code == "PLAN_STALE"


def test_plan_unknown_id_expired() -> None:
    with pytest.raises(McpError) as excinfo:
        get_plan("no-such-plan")
    assert excinfo.value.code == "PLAN_EXPIRED"
