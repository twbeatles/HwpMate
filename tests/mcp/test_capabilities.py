from __future__ import annotations

from hwpmate import mcp as mcp_package
from hwpmate.constants import FORMAT_TYPES, VERSION
from hwpmate.mcp import MCP_SCHEMA_VERSION
from hwpmate.mcp.capabilities import gather_capabilities, list_supported_formats
from hwpmate.mcp.config import McpConfig


def test_capabilities_schema_and_formats() -> None:
    caps = gather_capabilities(McpConfig())
    assert caps["schema_version"] == MCP_SCHEMA_VERSION
    assert caps["app_version"] == VERSION
    assert caps["hancom_com_probe"] == "not_run"
    assert caps["instance_lock"] == "unknown"
    assert set(caps["supported_formats"]) == set(FORMAT_TYPES.keys())
    assert caps["cli_path_valid"] is False  # 미설정 기본값은 fail closed 보고


def test_capabilities_do_not_claim_com_available() -> None:
    caps = gather_capabilities(McpConfig(cli_executable="C:/no/such/HwpMate.exe"))
    assert caps["cli_path_valid"] is False
    assert caps["hancom_com_probe"] == "not_run"
    assert any("현장 테스트" in str(w) for w in caps["warnings"])


def test_list_supported_formats_covers_all_format_types() -> None:
    result = list_supported_formats()
    assert result["schema_version"] == MCP_SCHEMA_VERSION
    names = [item["name"] for item in result["formats"]]
    assert set(names) == set(FORMAT_TYPES.keys())
    assert len(names) == len(FORMAT_TYPES)


def test_package_exports_schema_version() -> None:
    assert mcp_package.MCP_SCHEMA_VERSION == "hwpmate-mcp/v1"
