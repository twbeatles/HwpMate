"""변환 엔진 Protocol."""

from __future__ import annotations

from typing import Protocol


class ConverterEngine(Protocol):
    @property
    def progid_used(self) -> str | None: ...

    pdf_export_mode: str

    def initialize(self, *, manage_com_apartment: bool = True) -> bool: ...
    def convert_file(
        self,
        input_path,
        output_path,
        format_type="PDF",
        *,
        cancel_check=None,
    ) -> tuple[bool, str | None]: ...
    def cleanup(self) -> None: ...
    def has_owned_processes(self) -> bool: ...
    def kill_owned_processes(self) -> bool: ...
