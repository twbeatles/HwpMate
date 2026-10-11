"""P0-A/C — 공유 실행 경계(execute_task)의 원본·출력 보호 회귀 테스트.

GUI ConversionWorker 와 헤드리스 CLI 가 공유하는
hwpmate.workers.conversion_worker.task_runner.execute_task 를 직접 검증한다.
계획 시점 resolve_output_conflicts 를 우회·경과한 뒤에도 변환 직전에
출력 충돌을 다시 할당하는지, 원본 SHA-256 불변식을 지킨다.
"""

from __future__ import annotations

import hashlib
from pathlib import Path

from hwpmate.models import ConversionTask, PlannedConversion
from hwpmate.services.task_planner import TaskPlanner
from hwpmate.workers.conversion_worker.summary import build_summary
from hwpmate.workers.conversion_worker.task_runner import TaskRunOptions, execute_task


def _sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


class WritingConverter:
    """받은 output_path 에 그대로 쓰는 충실한 fake 엔진 (ConverterEngine 충족)."""

    def __init__(self, *, payload: bytes | None = None, succeed: bool = True) -> None:
        self.payload = payload if payload is not None else b"%PDF-1.4 fake\n"
        self.succeed = succeed
        self.seen_outputs: list[Path] = []
        self.progid_used: str | None = "Fake.Hwp"
        self.pdf_export_mode = "saveas_first"

    def initialize(self, *, manage_com_apartment: bool = True) -> bool:
        del manage_com_apartment
        return True

    def cleanup(self) -> None:
        return None

    def has_owned_processes(self) -> bool:
        return False

    def kill_owned_processes(self) -> bool:
        return False

    def convert_file(self, input_path, output_path, format_type="PDF", *, cancel_check=None):
        del input_path, format_type, cancel_check
        target = Path(output_path)
        self.seen_outputs.append(target)
        if self.succeed:
            target.write_bytes(self.payload_for(target))
        return self.succeed, None if self.succeed else "fake failure"

    def payload_for(self, target: Path) -> bytes:  # noqa: ARG002
        return self.payload

    @property
    def last_created_files(self) -> list[Path]:
        return list(self.seen_outputs) if self.succeed else []

    @property
    def last_output_size(self) -> int | None:
        if not self.succeed or not self.seen_outputs:
            return None
        try:
            return self.seen_outputs[-1].stat().st_size
        except OSError:
            return None


class PerInputConverter(WritingConverter):
    def payload_for(self, target: Path) -> bytes:
        return f"%PDF-1.4 from-{target.stem}\n".encode("ascii")


def _options(format_type: str = "PDF", overwrite: bool = False) -> TaskRunOptions:
    return TaskRunOptions(
        format_type=format_type,
        overwrite=overwrite,
        backup_enabled=False,
        retry_count=0,
    )


def test_runtime_toctou_planted_output_is_never_silently_replaced(tmp_path: Path) -> None:
    """계획 후 누군가 출력 파일을 만들었어도 무음으로 교체하지 않는다."""
    src = tmp_path / "a.hwp"
    src.write_bytes(b"dummy-hwp")
    out_dir = tmp_path / "out"
    out_dir.mkdir()
    planted = out_dir / "a.pdf"
    planted.write_bytes(b"someone was here first")
    planted_hash = _sha256(planted)

    task = ConversionTask(input_file=src, output_file=out_dir / "a.pdf")
    converter = WritingConverter(payload=b"%PDF-1.4 converted\n")
    warnings = execute_task(
        task,
        converter,
        _options(),
        planner=TaskPlanner(),
        used_output_path_keys=set(),
        cancel_check=lambda: False,
    )

    assert task.status == "성공"
    assert planted.read_bytes() == b"someone was here first"
    assert _sha256(planted) == planted_hash
    assert task.output_file != planted
    assert task.output_file.read_bytes() == b"%PDF-1.4 converted\n"
    assert converter.seen_outputs == [task.output_file]
    assert any("변환 직전" in warning for warning in warnings)


def test_runtime_overwrite_true_never_clobbers_protected_source(tmp_path: Path) -> None:
    """overwrite=true 여도 기존 원본 .hwp/.hwpx 위에는 절대 쓰지 않는다."""
    for source_name, target_format in (("a.hwp", "HWP"), ("a.hwpx", "HWPX")):
        work = tmp_path / source_name.replace(".", "_")
        work.mkdir()
        origin = work / source_name
        origin.write_bytes(b"original-document:" + source_name.encode())
        origin_hash = _sha256(origin)
        victim_in = work / f"other{'.hwpx' if source_name.endswith('.hwp') else '.hwp'}"
        victim_in.write_bytes(b"dummy-hwp")

        # 오래된 계획·직접 조작으로 출력이 원본과 겹친 최악 조건을 재현한다.
        task = ConversionTask(input_file=victim_in, output_file=origin)
        converter = WritingConverter(payload=b"converted-bytes\n")
        warnings = execute_task(
            task,
            converter,
            _options(format_type=target_format, overwrite=True),
            planner=TaskPlanner(),
            used_output_path_keys=set(),
            cancel_check=lambda: False,
        )

        assert task.status == "성공", source_name
        assert _sha256(origin) == origin_hash, source_name
        assert task.output_file != origin, source_name
        assert task.output_file.read_bytes() == b"converted-bytes\n", source_name
        assert any("변환 직전" in warning for warning in warnings), source_name


def test_runtime_batch_duplicates_do_not_intermingle(tmp_path: Path) -> None:
    """같은 출력 파일명을 요구한 두 입력은 분리 저장되고 내용이 섞이지 않는다."""
    in_dir = tmp_path / "in"
    out_dir = tmp_path / "out"
    in_dir.mkdir()
    out_dir.mkdir()
    first = in_dir / "a.hwp"
    second = in_dir / "b.hwp"
    first.write_bytes(b"dummy-a")
    second.write_bytes(b"dummy-b")

    tasks = [
        ConversionTask(input_file=first, output_file=out_dir / "same.pdf"),
        ConversionTask(input_file=second, output_file=out_dir / "same.pdf"),
    ]
    converter = PerInputConverter()
    used: set[str] = set()
    for task in tasks:
        warnings = execute_task(
            task,
            converter,
            _options(),
            planner=TaskPlanner(),
            used_output_path_keys=used,
            cancel_check=lambda: False,
        )
        assert task.status == "성공"
        assert warnings == [] or all("변환 직전" in warning for warning in warnings)

    assert tasks[0].output_file != tasks[1].output_file
    first_bytes = tasks[0].output_file.read_bytes()
    second_bytes = tasks[1].output_file.read_bytes()
    assert first_bytes != second_bytes
    assert b"same" in first_bytes and b"same" in second_bytes
    # 두 산출물 모두 디스크에 실재한다.
    assert tasks[0].output_file.is_file() and tasks[1].output_file.is_file()


def test_claimed_but_missing_artifact_is_failure_not_success(tmp_path: Path) -> None:
    """성공 반환 + 보고된 산출물 없음 = 실패 (COM 이후 삭제·경쟁 포함)."""
    src = tmp_path / "a.hwp"
    src.write_bytes(b"dummy-hwp")
    task = ConversionTask(input_file=src, output_file=tmp_path / "a.pdf")

    class LyingConverter(WritingConverter):
        def convert_file(self, input_path, output_path, format_type="PDF", *, cancel_check=None):  # noqa: ANN001, ANN202
            del input_path, output_path, format_type, cancel_check
            self.seen_outputs.append(Path(tmp_path / "a.pdf"))
            return True, None

    execute_task(
        task,
        LyingConverter(),
        _options(),
        planner=TaskPlanner(),
        used_output_path_keys=set(),
        cancel_check=lambda: False,
    )
    assert task.status == "실패"
    assert task.error is not None and "산출물" in task.error


def test_claimed_zero_byte_artifact_is_failure_not_success(tmp_path: Path) -> None:
    """성공 반환 + 0바이트 산출물 = 실패."""
    src = tmp_path / "a.hwp"
    src.write_bytes(b"dummy-hwp")
    task = ConversionTask(input_file=src, output_file=tmp_path / "a.pdf")
    execute_task(
        task,
        WritingConverter(payload=b""),
        _options(),
        planner=TaskPlanner(),
        used_output_path_keys=set(),
        cancel_check=lambda: False,
    )
    assert task.status == "실패"
    assert task.error is not None and "비어" in task.error


def test_summary_counts_match_files_on_disk(tmp_path: Path) -> None:
    """summary 성공 집계와 파일시스템 실재 목록을 대조한다."""
    in_dir = tmp_path / "in"
    out_dir = tmp_path / "out"
    in_dir.mkdir()
    out_dir.mkdir()
    good = in_dir / "good.hwp"
    bad = in_dir / "bad.hwp"
    skipped = in_dir / "same.pdf"
    good.write_bytes(b"dummy")
    bad.write_bytes(b"dummy")
    skipped.write_bytes(b"dummy")

    tasks = [
        ConversionTask(input_file=good, output_file=out_dir / "good.pdf"),
        ConversionTask(input_file=bad, output_file=out_dir / "bad.pdf"),
    ]
    results = {"good.hwp": (True, None), "bad.hwp": (False, "save failed")}

    class SplitConverter(WritingConverter):
        def convert_file(self, input_path, output_path, format_type="PDF", *, cancel_check=None):  # noqa: ANN001, ANN202
            del format_type, cancel_check
            ok, error = results[Path(input_path).name]
            if ok:
                Path(output_path).write_bytes(b"%PDF-1.4 ok\n")
                self.seen_outputs.append(Path(output_path))
            return ok, error

    converter = SplitConverter()
    used: set[str] = set()
    for task in tasks:
        execute_task(
            task,
            converter,
            _options(),
            planner=TaskPlanner(),
            used_output_path_keys=used,
            cancel_check=lambda: False,
        )
    planned = PlannedConversion(
        format_type="PDF",
        same_location=False,
        output_path=str(out_dir),
        tasks=tasks,
        skipped_tasks=[ConversionTask(skipped, skipped, status="건너뜀", error="same format")],
    )
    summary = build_summary(
        format_type="PDF",
        tasks=tasks,
        planned=planned,
        warnings=[],
        elapsed_seconds=0.1,
        progid_used="Fake.Hwp",
    )
    assert (summary.success_count, summary.failed_count, summary.skipped_count) == (1, 1, 1)
    for output in summary.output_paths:
        assert Path(output).is_file(), output
    assert len(summary.output_paths) == summary.success_count
