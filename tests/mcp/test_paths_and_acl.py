from __future__ import annotations

import os
from pathlib import Path

import pytest

from hwpmate.mcp.config import McpConfig
from hwpmate.mcp.errors import McpError
from hwpmate.mcp.policy import (
    check_input_path,
    check_output_dir,
    normalize_format,
)


def _cfg(tmp_path: Path, **overrides: object) -> McpConfig:
    roots: dict[str, object] = {
        "input_roots": [str(tmp_path / "in")],
        "output_roots": [str(tmp_path / "out")],
    }
    roots.update(overrides)
    return McpConfig.from_mapping(roots)


def test_no_roots_fail_closed(tmp_path: Path) -> None:
    cfg = McpConfig()
    with pytest.raises(McpError) as excinfo:
        check_input_path(cfg, str(tmp_path))
    assert excinfo.value.code == "INPUT_DENIED"
    with pytest.raises(McpError) as excinfo:
        check_output_dir(cfg, str(tmp_path))
    assert excinfo.value.code == "OUTPUT_DENIED"


def test_outside_root_denied(tmp_path: Path) -> None:
    cfg = _cfg(tmp_path)
    (tmp_path / "in").mkdir()
    outside = tmp_path / "elsewhere"
    outside.mkdir()
    with pytest.raises(McpError) as excinfo:
        check_input_path(cfg, str(outside))
    assert excinfo.value.code == "INPUT_DENIED"


def test_dotdot_escape_denied(tmp_path: Path) -> None:
    cfg = _cfg(tmp_path)
    (tmp_path / "in").mkdir()
    (tmp_path / "secret.hwp").write_bytes(b"x")
    with pytest.raises(McpError) as excinfo:
        check_input_path(cfg, str(tmp_path / "in" / ".." / "secret.hwp"))
    assert excinfo.value.code == "INPUT_DENIED"


def test_unc_denied_by_default(tmp_path: Path) -> None:
    cfg = _cfg(tmp_path)
    with pytest.raises(McpError) as excinfo:
        check_input_path(cfg, r"\\server\share\doc.hwp")
    assert excinfo.value.code == "INPUT_DENIED"


def test_symlink_denied_by_default(tmp_path: Path) -> None:
    allowed = tmp_path / "in"
    allowed.mkdir()
    outside = tmp_path / "real"
    outside.mkdir()
    (outside / "doc.hwp").write_bytes(b"x")
    link = allowed / "evil"
    try:
        link.symlink_to(outside, target_is_directory=True)
    except OSError:
        pytest.skip("symlink 생성 권한 없음")
    cfg = _cfg(tmp_path)
    with pytest.raises(McpError) as excinfo:
        check_input_path(cfg, str(link / "doc.hwp"))
    assert excinfo.value.code == "INPUT_DENIED"


def test_inside_root_allowed(tmp_path: Path) -> None:
    cfg = _cfg(tmp_path)
    target = tmp_path / "in" / "doc.hwp"
    target.parent.mkdir(parents=True)
    target.write_bytes(b"x")
    assert check_input_path(cfg, str(target)).name == "doc.hwp"


def test_bad_format_rejected() -> None:
    with pytest.raises(McpError) as excinfo:
        normalize_format("EXE")
    assert excinfo.value.code == "BAD_FORMAT"
    assert normalize_format("pdf") == "PDF"


def test_output_requires_explicit_dir(tmp_path: Path) -> None:
    from hwpmate.mcp.planner_adapter import preview_conversion

    cfg = _cfg(tmp_path)
    src = tmp_path / "in" / "a.hwp"
    src.parent.mkdir(parents=True)
    src.write_bytes(b"x")
    with pytest.raises(McpError) as excinfo:
        preview_conversion(cfg, input_paths=[str(src)], format_type="PDF", output_dir="")
    assert excinfo.value.code == "OUTPUT_DENIED"
    _ = os  # unused guard
