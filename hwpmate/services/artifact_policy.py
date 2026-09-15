from __future__ import annotations

import os
import re
from pathlib import Path

from ..constants import BACKUP_DIR_NAME, SUPPORTED_EXTENSIONS

AUXILIARY_ARTIFACT_FORMATS = frozenset({"HTML", "PNG", "JPG", "BMP", "GIF"})
IMAGE_ARTIFACT_FORMATS = frozenset({"PNG", "JPG", "BMP", "GIF"})
# 공백·괄호는 제외: "report (1)" 같은 형제 문서의 산출물을 "report" 의 보조 산출물로 오인한다.
AUXILIARY_NAME_DELIMITERS = frozenset({"_", "-", "."})
MAX_AUXILIARY_SCAN_FILES = 500

# 한글 2022(12.0) 실측: 이미지 SaveAs 는 요청 경로 대신 페이지마다
# "{stem}001.png", "{stem}002.png" 처럼 구분자 없이 3자리 번호를 붙여 저장한다.
# 자릿수를 넓히면 형제 문서 "report1.hwp" 의 "report1001.png" 를 "report" 의 페이지로 오인한다.
_PAGE_NUMBER_REST = re.compile(r"^\d{3}(\.|$)")
_IMAGE_PAGE_SUFFIX = re.compile(r"^[_-]?\d{3}$")
# 한글 2022 HTML 저장 시 본문 이미지·수식은 문서 이름과 무관한 "PIC388B.png" 형태로 생성된다.
_HTML_EMBEDDED_IMAGE = re.compile(r"^PIC[0-9A-F]+\.(png|jpe?g|gif|bmp|wmf|emf)$", re.IGNORECASE)

# 산출물 후보에서 제외: 원본 한글 문서는 절대 산출물로 취급하지 않는다.
PROTECTED_SOURCE_EXTENSIONS = frozenset(ext.lower() for ext in SUPPORTED_EXTENSIONS)
EXCLUDED_ARTIFACT_DIR_NAMES = frozenset({BACKUP_DIR_NAME.lower()})


def uses_auxiliary_artifacts(format_type: str) -> bool:
    return format_type.upper() in AUXILIARY_ARTIFACT_FORMATS


def matches_artifact_stem(name: str, stem: str, *, allow_page_number: bool = False) -> bool:
    """Return True when a file/directory name belongs to the output stem."""
    name_key = name.lower()
    stem_key = stem.lower()
    if not stem_key:
        return False
    if name_key == stem_key:
        return True
    if not name_key.startswith(stem_key):
        return False
    if len(name_key) == len(stem_key):
        return True
    rest = name_key[len(stem_key):]
    if rest[0] in AUXILIARY_NAME_DELIMITERS:
        return True
    return allow_page_number and bool(_PAGE_NUMBER_REST.match(rest))


def is_protected_source_path(path: Path) -> bool:
    """변환 입력이 될 수 있는 원본 한글 문서 경로인지 (확장자 기준)."""
    return path.suffix.lower() in PROTECTED_SOURCE_EXTENSIONS


def is_artifact_candidate(
    path: Path,
    output_file: Path,
    format_type: str,
    *,
    for_conflict: bool = False,
) -> bool:
    """output_file 변환의 산출물(기본/보조)로 볼 수 있는 경로인지.

    - 기본 출력 파일 이름은 항상 후보
    - 원본 한글 문서(.hwp/.hwpx)와 앱 백업 폴더는 제외
    - 이미지 형식은 같은 확장자의 페이지 파일만 ({stem}001, {stem}_001)
    - HTML 의 문서 이름과 무관한 PIC* 이미지는 변경 감지(스냅샷)용 후보일 뿐 충돌로 보지 않는다
    """
    name = path.name
    if name.lower() == output_file.name.lower():
        return True
    if not uses_auxiliary_artifacts(format_type):
        return False
    fmt = format_type.upper()
    try:
        is_dir = path.is_dir()
    except OSError:
        return False
    if is_dir:
        if name.lower() in EXCLUDED_ARTIFACT_DIR_NAMES:
            return False
        return matches_artifact_stem(name, output_file.stem)
    if is_protected_source_path(path):
        return False
    if fmt in IMAGE_ARTIFACT_FORMATS:
        if path.suffix.lower() != output_file.suffix.lower():
            return False
        base_key = path.stem.lower()
        stem_key = output_file.stem.lower()
        if not stem_key or not base_key.startswith(stem_key):
            return False
        return bool(_IMAGE_PAGE_SUFFIX.match(base_key[len(stem_key):]))
    if fmt == "HTML" and not for_conflict and _HTML_EMBEDDED_IMAGE.match(name):
        return True
    return matches_artifact_stem(name, output_file.stem)


def artifact_key(path: Path) -> str:
    return os.path.normcase(str(path.resolve() if path.exists() else path.absolute()))


def iter_candidate_artifact_paths(
    output_file: Path,
    format_type: str,
    *,
    include_nested: bool = True,
    nested_limit: int = MAX_AUXILIARY_SCAN_FILES,
) -> list[Path]:
    candidates: dict[str, Path] = {artifact_key(output_file): output_file}
    if not uses_auxiliary_artifacts(format_type):
        return list(candidates.values())

    parent = output_file.parent
    if not parent.exists():
        return list(candidates.values())

    nested_count = 0
    try:
        for child in parent.iterdir():
            if not is_artifact_candidate(child, output_file, format_type):
                continue
            if child.is_file():
                candidates[artifact_key(child)] = child
                continue
            if child.is_dir() and include_nested:
                if nested_count >= nested_limit:
                    continue
                try:
                    for nested in child.rglob("*"):
                        if not nested.is_file():
                            continue
                        if nested_count >= nested_limit:
                            break
                        candidates[artifact_key(nested)] = nested
                        nested_count += 1
                except OSError:
                    continue
    except OSError:
        return list(candidates.values())

    return list(candidates.values())


def existing_artifact_conflicts(output_file: Path, format_type: str) -> list[Path]:
    conflicts: dict[str, Path] = {}
    if output_file.exists():
        conflicts[artifact_key(output_file)] = output_file

    if not uses_auxiliary_artifacts(format_type):
        return list(conflicts.values())

    parent = output_file.parent
    if not parent.exists():
        return list(conflicts.values())

    try:
        for child in parent.iterdir():
            if child == output_file:
                continue
            if is_artifact_candidate(child, output_file, format_type, for_conflict=True):
                conflicts[artifact_key(child)] = child
    except OSError:
        return list(conflicts.values())

    return sorted(conflicts.values(), key=lambda path: str(path).lower())
