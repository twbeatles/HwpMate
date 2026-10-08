"""산출물 검증 (구현 설계서 Phase 3).

job 결과에 기록된 출력 파일의 존재·크기·확장자·매직 서명을 확인한다.
원본과의 시각 동일성은 보증하지 않으며, PDF 페이지 수 검사는 하지 않는다.
읽기 전용이다.
"""

from __future__ import annotations

from pathlib import Path

from ..constants import FORMAT_TYPES
from .config import McpConfig
from .errors import McpError
from .jobs import TERMINAL_STATES, JobManager
from .schemas import ok_envelope

_IMAGE_PAGE_SUFFIXES = (".png", ".jpg", ".jpeg", ".bmp", ".gif")


def _signature_ok(format_type: str, head: bytes, text_head: str) -> tuple[bool, str]:
    fmt = format_type.upper()
    if fmt == "PDF":
        return (head.startswith(b"%PDF-"), "PDF 서명(%PDF-) 없음")
    if fmt in ("DOCX", "ODT", "HWPX"):
        return (head.startswith(b"PK\x03\x04"), "ZIP 컨테이너 서명 없음")
    if fmt == "HWP":
        return (head.startswith(b"\xd0\xcf\x11\xe0"), "OLE 문서 서명 없음")
    if fmt == "PNG":
        return (head.startswith(b"\x89PNG"), "PNG 서명 없음")
    if fmt == "JPG":
        return (head.startswith(b"\xff\xd8\xff"), "JPEG 서명 없음")
    if fmt == "BMP":
        return (head.startswith(b"BM"), "BMP 서명 없음")
    if fmt == "GIF":
        return (head.startswith(b"GIF8"), "GIF 서명 없음")
    if fmt == "RTF":
        return (text_head.lstrip().startswith("{\\rtf"), "RTF 헤더 없음")
    if fmt == "HTML":
        lowered = text_head.lower().lstrip()
        return (
            lowered.startswith(("<!doctype html", "<html")),
            "HTML 문서 구조를 확인하지 못함(경고 수준)",
        )
    if fmt == "TXT":
        return (True, "")
    return (True, "")


def _expected_ext(format_type: str) -> str:
    return str(FORMAT_TYPES[format_type.upper()]["ext"]).lower()


def validate_job_outputs(
    manager: JobManager, job_id: str, *, cfg: McpConfig | None = None
) -> dict:
    del cfg
    record = manager.get(job_id)
    if record.state not in TERMINAL_STATES:
        raise McpError("JOB_NOT_DONE", f"아직 끝나지 않은 작업은 검증할 수 없습니다: {record.state}")
    expected_ext = _expected_ext(record.format)
    files: list[dict] = []
    for item in record.items:
        if str(item.get("status", "")) != "성공":
            continue
        for output in item.get("outputs", []) or []:
            path = Path(str(output))
            entry: dict = {"path": str(output), "verified": False, "reasons": []}
            try:
                if not path.is_file():
                    entry["reasons"].append("파일 없음")
                    files.append(entry)
                    continue
                size = path.stat().st_size
                if size <= 0:
                    entry["reasons"].append("빈 파일(0 bytes)")
                    files.append(entry)
                    continue
                if path.suffix.lower() != expected_ext and not (
                    record.format.upper() == "JPG" and path.suffix.lower() == ".jpeg"
                ):
                    entry["reasons"].append(f"확장자 불일치(기대 {expected_ext})")
                    files.append(entry)
                    continue
                try:
                    head = path.read_bytes()[:16]
                except OSError as exc:
                    entry["reasons"].append(f"읽기 실패: {exc}")
                    files.append(entry)
                    continue
                try:
                    text_head = head.decode("utf-8", errors="replace")[:64]
                except Exception:
                    text_head = ""
                ok, reason = _signature_ok(record.format, head, text_head)
                if not ok:
                    entry["reasons"].append(reason)
                    files.append(entry)
                    continue
                entry["verified"] = True
                entry["size"] = size
                files.append(entry)
            except OSError as exc:
                entry["reasons"].append(f"검사 실패: {exc}")
                files.append(entry)
    invalid = [entry for entry in files if not entry["verified"]]
    return ok_envelope(
        job_id=record.job_id,
        state=record.state,
        format=record.format,
        checked=len(files),
        verified=len(files) - len(invalid),
        invalid=invalid,
        all_verified=not invalid,
    )
