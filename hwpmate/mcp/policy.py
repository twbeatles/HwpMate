"""경로·포맷 허용 정책 (구현 설계서 §6.1).

- 모든 경로는 resolve() 후 root containment를 검사한다.
- 루트 설정이 없으면 변환 관련 경로는 fail closed 한다.
- UNC·reparse point(symlink/junction)는 기본 거부한다.
"""

from __future__ import annotations

import os
from pathlib import Path

from ..constants import FORMAT_TYPES
from .config import McpConfig
from .errors import McpError


def normalize_format(format_type: str) -> str:
    name = str(format_type or "").upper().strip()
    if name not in FORMAT_TYPES:
        raise McpError(
            "BAD_FORMAT",
            f"지원하지 않는 출력 형식: {format_type} "
            f"(가능한 형식: {', '.join(FORMAT_TYPES.keys())})",
        )
    return name


def _is_unc_path(raw: str) -> bool:
    text = str(raw).replace("/", "\\")
    return text.startswith("\\\\")


def _has_reparse_point(path: Path) -> bool:
    """경로 구성 요소에 symlink/junction이 있으면 True (best-effort)."""
    current = path
    if not current.is_absolute():
        current = Path(os.path.abspath(str(current)))
    parts = list(current.parts)
    for index in range(1, len(parts) + 1):
        candidate = Path(*parts[:index])
        try:
            if candidate.is_symlink():
                return True
            if os.name == "nt":
                try:
                    attrs = os.stat(candidate, follow_symlinks=False).st_file_attributes  # type: ignore[attr-defined]
                    if attrs & 0x400:  # FILE_ATTRIBUTE_REPARSE_POINT
                        return True
                except (OSError, AttributeError):
                    pass
        except OSError:
            return False
    return False


def resolve_absolute(raw: str | Path, *, what: str = "경로") -> Path:
    """symlink를 해석하지 않은 절대 경로 (reparse 검사용)."""
    text = str(raw or "").strip()
    if not text:
        raise McpError("INPUT_DENIED", f"{what}이 비어 있습니다.")
    if "\x00" in text:
        raise McpError("INPUT_DENIED", f"{what}에 NUL 문자를 사용할 수 없습니다.")
    return Path(os.path.abspath(os.path.expanduser(text)))


def resolve_strict(raw: str | Path, *, what: str = "경로") -> Path:
    absolute = resolve_absolute(raw, what=what)
    check_reparse_on_absolute(absolute)
    try:
        return absolute.resolve()
    except OSError as exc:
        raise McpError("INPUT_DENIED", f"{what} 해석 실패: {exc}") from exc


def check_reparse_on_absolute(absolute: Path) -> None:
    if _has_reparse_point(absolute):
        raise McpError("INPUT_DENIED", f"reparse point(symlink/junction)를 포함해 거부합니다: {absolute}")


def check_containment(resolved: Path, roots: list[str], *, kind: str) -> Path:
    if not roots:
        code = "INPUT_DENIED" if kind == "입력" else "OUTPUT_DENIED"
        raise McpError(code, f"{kind} 허용 루트가 설정되지 않아 {kind} 경로를 거부합니다.")
    for root in roots:
        try:
            anchor = Path(os.path.abspath(os.path.expanduser(str(root)))).resolve()
        except OSError:
            continue
        try:
            resolved.relative_to(anchor)
            return resolved
        except ValueError:
            continue
    code = "INPUT_DENIED" if kind == "입력" else "OUTPUT_DENIED"
    raise McpError(code, f"{kind} 허용 범위를 벗어난 경로입니다: {resolved}")


def _resolve_with_policy(cfg: McpConfig, raw: str | Path, *, what: str) -> Path:
    absolute = resolve_absolute(raw, what=what)
    if not cfg.allow_reparse_points:
        try:
            check_reparse_on_absolute(absolute)
        except OSError:
            pass
    try:
        return absolute.resolve()
    except OSError as exc:
        raise McpError("INPUT_DENIED", f"{what} 해석 실패: {exc}") from exc


def check_input_path(cfg: McpConfig, raw: str | Path) -> Path:
    if _is_unc_path(str(raw)) and not cfg.allow_unc_paths:
        raise McpError("INPUT_DENIED", f"UNC 경로는 허용되지 않습니다: {raw}")
    resolved = _resolve_with_policy(cfg, raw, what="입력 경로")
    return check_containment(resolved, cfg.input_roots, kind="입력")


def check_output_dir(cfg: McpConfig, raw: str | Path) -> Path:
    if _is_unc_path(str(raw)) and not cfg.allow_unc_paths:
        raise McpError("OUTPUT_DENIED", f"UNC 경로는 허용되지 않습니다: {raw}")
    resolved = _resolve_with_policy(cfg, raw, what="출력 폴더")
    return check_containment(resolved, cfg.output_roots, kind="출력")


def is_within_directory(candidate: Path, directory: Path) -> bool:
    try:
        candidate.relative_to(directory)
        return True
    except ValueError:
        return False
