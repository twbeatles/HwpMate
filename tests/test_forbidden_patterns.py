"""CLAUDE.md §2 '절대 깨지면 안 되는 로직'의 금지 호출을 정적으로 고정한다.

- 물리 인쇄: HAction/CreateAction Execute 는 PrintToPDFEx 외 금지
- 문서 닫기: Clear 는 1(버림)만 허용 (2/3 은 원본을 디스크에 저장)
- 강제 종료: taskkill 은 PID 지정만 허용 (이미지명 /IM 전역 종료 금지)
"""

from __future__ import annotations

import re
from pathlib import Path

PACKAGE = Path(__file__).resolve().parents[1] / "hwpmate"


def _sources() -> list[tuple[Path, str]]:
    return [(path, path.read_text(encoding="utf-8")) for path in sorted(PACKAGE.rglob("*.py"))]


def _code_lines(text: str) -> list[str]:
    return [line.split("#", 1)[0] for line in text.splitlines()]


def test_execute_only_print_to_pdf_ex() -> None:
    offenders = []
    for path, text in _sources():
        for line in _code_lines(text):
            for match in re.finditer(r"\.Execute\(\s*([^,)]*)", line):
                if match.group(1).strip() != '"PrintToPDFEx"':
                    offenders.append(f"{path.name}: {line.strip()}")
    assert offenders == []


def test_clear_only_discard_option() -> None:
    offenders = []
    for path, text in _sources():
        for line in _code_lines(text):
            for match in re.finditer(r"\.Clear\(\s*([^)]*)\)", line):
                argument = match.group(1).replace(" ", "")
                if argument not in {"1", "option=1"}:
                    offenders.append(f"{path.name}: {line.strip()}")
    assert offenders == []


def test_taskkill_targets_pid_only() -> None:
    offenders = []
    for path, text in _sources():
        code = "\n".join(_code_lines(text))
        for match in re.finditer(r"\[\s*\"taskkill\"[^\]]*\]", code, flags=re.DOTALL):
            call = match.group(0)
            if '"/IM"' in call.upper() or '"/PID"' not in call:
                offenders.append(f"{path.name}: {call}")
    assert offenders == []
