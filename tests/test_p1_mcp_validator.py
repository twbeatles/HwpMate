"""P1-B — MCP 산출물 검증기(validator) 단위 테스트.

hwpmate/mcp/validator.py 커버리지 공백(52%)을 메운다.
정상 exit/report 라도 파일 없음·0바이트·확장자 불일치·매직 서명 불일치를
검증 실패로 구분하는지, 스텁 job 레코드로 직접 증명한다.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path

import pytest

from hwpmate.mcp.errors import McpError
from hwpmate.mcp.validator import validate_job_outputs


@dataclass
class _StubRecord:
    job_id: str
    state: str
    format: str
    items: list[dict] = field(default_factory=list)


@dataclass
class _StubManager:
    record: _StubRecord

    def get(self, job_id: str) -> _StubRecord:
        assert job_id == self.record.job_id
        return self.record


def _validate(tmp_path: Path, fmt: str, payload: bytes | None, *, state: str = "succeeded") -> dict:
    target = tmp_path / f"out.{'jpeg' if fmt == 'JPG' else fmt.lower()}"
    items: list[dict] = []
    if payload is not None:
        target.write_bytes(payload)
        items = [{"status": "성공", "outputs": [str(target)]}]
    record = _StubRecord(job_id="j1", state=state, format=fmt, items=items)
    return validate_job_outputs(_StubManager(record), "j1")  # type: ignore[arg-type]


def test_validate_accepts_real_pdf_signature(tmp_path: Path) -> None:
    result = _validate(tmp_path, "PDF", b"%PDF-1.4\n1 0 obj\n")
    assert result["all_verified"] is True
    assert result["checked"] == 1 and result["verified"] == 1


def test_validate_accepts_docx_zip_signature(tmp_path: Path) -> None:
    result = _validate(tmp_path, "DOCX", b"PK\x03\x04" + b"\x00" * 32)
    assert result["all_verified"] is True


def test_validate_accepts_hwp_ole_signature(tmp_path: Path) -> None:
    result = _validate(tmp_path, "HWP", b"\xd0\xcf\x11\xe0" + b"\x00" * 32)
    assert result["all_verified"] is True


def test_validate_accepts_jpg_with_jpeg_extension(tmp_path: Path) -> None:
    result = _validate(tmp_path, "JPG", b"\xff\xd8\xff" + b"\x00" * 16)
    assert result["all_verified"] is True


def test_validate_rejects_missing_file(tmp_path: Path) -> None:
    record = _StubRecord(
        job_id="j1",
        state="succeeded",
        format="PDF",
        items=[{"status": "성공", "outputs": [str(tmp_path / "gone.pdf")]}],
    )
    result = validate_job_outputs(_StubManager(record), "j1")  # type: ignore[arg-type]
    assert result["all_verified"] is False
    assert result["invalid"][0]["reasons"] == ["파일 없음"]


def test_validate_rejects_zero_byte_file(tmp_path: Path) -> None:
    result = _validate(tmp_path, "PDF", b"")
    assert result["all_verified"] is False
    assert result["invalid"][0]["reasons"] == ["빈 파일(0 bytes)"]


def test_validate_rejects_wrong_magic_for_pdf(tmp_path: Path) -> None:
    """exit 0 + 성공 report 라도 PDF 매직이 없으면 실패로 구분한다."""
    result = _validate(tmp_path, "PDF", b"PK\x03\x04" + b"\x00" * 32)
    assert result["all_verified"] is False
    assert any("PDF" in reason for reason in result["invalid"][0]["reasons"])


def test_validate_rejects_extension_mismatch(tmp_path: Path) -> None:
    target = tmp_path / "out.txt"
    target.write_bytes(b"%PDF-1.4\n")
    record = _StubRecord(
        job_id="j1",
        state="succeeded",
        format="PDF",
        items=[{"status": "성공", "outputs": [str(target)]}],
    )
    result = validate_job_outputs(_StubManager(record), "j1")  # type: ignore[arg-type]
    assert result["all_verified"] is False
    assert any("확장자" in reason for reason in result["invalid"][0]["reasons"])


def test_validate_skips_failed_items(tmp_path: Path) -> None:
    target = tmp_path / "out.pdf"
    target.write_bytes(b"corrupt")
    record = _StubRecord(
        job_id="j1",
        state="partially_failed",
        format="PDF",
        items=[{"status": "실패", "outputs": [str(target)]}],
    )
    result = validate_job_outputs(_StubManager(record), "j1")  # type: ignore[arg-type]
    assert result["checked"] == 0
    assert result["all_verified"] is True


def test_validate_refuses_unfinished_job(tmp_path: Path) -> None:
    with pytest.raises(McpError) as excinfo:
        _validate(tmp_path, "PDF", b"%PDF-1.4\n", state="running")
    assert excinfo.value.code == "JOB_NOT_DONE"


def test_validate_txt_passes_without_signature(tmp_path: Path) -> None:
    result = _validate(tmp_path, "TXT", "한글 본문 그대로\n".encode("utf-8"))
    assert result["all_verified"] is True
