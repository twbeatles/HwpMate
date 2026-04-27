from __future__ import annotations

from datetime import datetime as real_datetime
from pathlib import Path

from hwpmate.models import ConversionTask, PlannedConversion
from hwpmate.workers.conversion_worker import ConversionWorker


class StubConverter:
    def __init__(self, *, results: dict[str, tuple[bool, str | None]], owned: bool = True, on_convert=None) -> None:
        self.results = results
        self.owned = owned
        self.on_convert = on_convert
        self.progid_used = "Stub.Hwp"
        self.cleaned = False
        self.kill_called = False

    def initialize(self) -> bool:
        return True

    def convert_file(self, input_path, output_path, format_type="PDF"):
        if self.on_convert is not None:
            self.on_convert(Path(input_path))
        return self.results[Path(input_path).name]

    def cleanup(self) -> None:
        self.cleaned = True

    def has_owned_processes(self) -> bool:
        return self.owned

    def kill_owned_processes(self) -> bool:
        self.kill_called = True
        return self.owned


class FailingInitConverter(StubConverter):
    def initialize(self) -> bool:
        raise RuntimeError("init failed")


class SequenceConverter(StubConverter):
    def __init__(self, *, sequence: list[tuple[bool, str | None]]) -> None:
        super().__init__(results={})
        self.sequence = sequence

    def convert_file(self, input_path, output_path, format_type="PDF"):
        del input_path, output_path, format_type
        return self.sequence.pop(0)


def test_conversion_worker_builds_summary_with_success_failure_and_skip(tmp_path: Path, monkeypatch) -> None:
    import hwpmate.workers.conversion_worker as worker_module

    monkeypatch.setattr(worker_module.time, "sleep", lambda _: None)
    first = tmp_path / "a.hwp"
    second = tmp_path / "b.hwp"
    skipped = tmp_path / "c.hwpx"
    for path in (first, second, skipped):
        path.write_text("x", encoding="utf-8")

    plan = PlannedConversion(
        format_type="PDF",
        same_location=True,
        output_path="",
        tasks=[
            ConversionTask(first, first.with_suffix(".pdf")),
            ConversionTask(second, second.with_suffix(".pdf")),
        ],
        skipped_tasks=[
            ConversionTask(skipped, skipped, status="건너뜀", error="이미 HWPX 형식입니다."),
        ],
        warnings=["동일 형식 1개는 자동으로 건너뜁니다."],
    )
    summaries = []
    worker = ConversionWorker(
        plan,
        converter_factory=lambda: StubConverter(
            results={
                "a.hwp": (True, None),
                "b.hwp": (False, "save failed"),
            }
        ),
    )
    worker.task_completed.connect(lambda summary: summaries.append(summary))
    worker.run()

    assert len(summaries) == 1
    summary = summaries[0]
    assert summary.success_count == 1
    assert summary.failed_count == 1
    assert summary.skipped_count == 1
    assert summary.canceled_count == 0
    assert summary.output_paths == [str(first.with_suffix(".pdf"))]
    failed = next(task for task in summary.tasks if task.status == "실패")
    assert failed.retry_count == 1


def test_conversion_worker_marks_remaining_tasks_as_canceled(tmp_path: Path) -> None:
    first = tmp_path / "a.hwp"
    second = tmp_path / "b.hwp"
    skipped = tmp_path / "c.hwp"
    for path in (first, second, skipped):
        path.write_text("x", encoding="utf-8")

    plan = PlannedConversion(
        format_type="PDF",
        same_location=True,
        output_path="",
        tasks=[
            ConversionTask(first, first.with_suffix(".pdf")),
            ConversionTask(second, second.with_suffix(".pdf")),
        ],
        skipped_tasks=[
            ConversionTask(skipped, skipped, status="건너뜀", error="이미 PDF 형식입니다."),
        ],
    )
    summaries = []
    worker = ConversionWorker(
        plan,
        converter_factory=lambda: StubConverter(
            results={
                "a.hwp": (True, None),
                "b.hwp": (True, None),
            },
            on_convert=lambda _: worker.cancel(),
        ),
    )
    worker.task_completed.connect(lambda summary: summaries.append(summary))
    worker.run()

    summary = summaries[0]
    assert summary.success_count == 1
    assert summary.canceled_count == 1
    assert summary.skipped_count == 1
    assert any(task.status == "취소됨" for task in summary.tasks)


def test_conversion_worker_emits_failed_summary_when_initialize_fails(tmp_path: Path) -> None:
    input_file = tmp_path / "a.hwp"
    input_file.write_text("x", encoding="utf-8")
    skipped = tmp_path / "b.hwpx"
    skipped.write_text("x", encoding="utf-8")
    plan = PlannedConversion(
        format_type="PDF",
        same_location=True,
        output_path="",
        tasks=[ConversionTask(input_file, input_file.with_suffix(".pdf"))],
        skipped_tasks=[ConversionTask(skipped, skipped, status="건너뜀", error="이미 HWPX 형식입니다.")],
    )
    summaries = []
    errors = []
    worker = ConversionWorker(
        plan,
        converter_factory=lambda: FailingInitConverter(results={"a.hwp": (True, None)}),
    )
    worker.task_completed.connect(lambda summary: summaries.append(summary))
    worker.error_occurred.connect(lambda error: errors.append(error))

    worker.run()

    assert errors == []
    assert len(summaries) == 1
    summary = summaries[0]
    assert summary.failed_count == 1
    assert summary.skipped_count == 1
    assert "한글 초기화 실패" in summary.failed_tasks[0].detail


def test_conversion_worker_retries_failed_conversion(tmp_path: Path, monkeypatch) -> None:
    import hwpmate.workers.conversion_worker as worker_module

    monkeypatch.setattr(worker_module.time, "sleep", lambda _: None)
    input_file = tmp_path / "a.hwp"
    input_file.write_text("x", encoding="utf-8")
    plan = PlannedConversion(
        format_type="PDF",
        same_location=True,
        output_path="",
        retry_count=1,
        tasks=[ConversionTask(input_file, input_file.with_suffix(".pdf"))],
    )
    summaries = []
    worker = ConversionWorker(
        plan,
        converter_factory=lambda: SequenceConverter(sequence=[(False, "temporary"), (True, None)]),
    )
    worker.task_completed.connect(lambda summary: summaries.append(summary))

    worker.run()

    task = summaries[0].tasks[0]
    assert task.status == "성공"
    assert task.retry_count == 1


def test_force_terminate_uses_owned_processes_only(tmp_path: Path) -> None:
    input_file = tmp_path / "a.hwp"
    input_file.write_text("x", encoding="utf-8")
    plan = PlannedConversion(
        format_type="PDF",
        same_location=True,
        output_path="",
        tasks=[ConversionTask(input_file, input_file.with_suffix(".pdf"))],
    )
    owned_converter = StubConverter(results={"a.hwp": (True, None)}, owned=True)
    worker = ConversionWorker(plan, converter_factory=lambda: owned_converter)
    worker.converter = owned_converter

    assert worker.can_force_terminate() is True
    assert worker.force_terminate() is True
    assert owned_converter.kill_called is True

    unowned_converter = StubConverter(results={"a.hwp": (True, None)}, owned=False)
    worker.converter = unowned_converter
    assert worker.can_force_terminate() is False
    assert worker.force_terminate() is False


def test_create_backup_avoids_name_collisions(tmp_path: Path, monkeypatch) -> None:
    source = tmp_path / "doc.hwp"
    source.write_text("x", encoding="utf-8")
    plan = PlannedConversion(format_type="PDF", same_location=True, output_path="")
    worker = ConversionWorker(plan)

    class FrozenDateTime:
        @classmethod
        def now(cls):
            return real_datetime(2026, 3, 18, 12, 0, 0, 123456)

    import hwpmate.workers.conversion_worker as worker_module

    monkeypatch.setattr(worker_module, "datetime", FrozenDateTime)
    worker._create_backup(source)
    worker._create_backup(source)

    backups = sorted((tmp_path / "backup").iterdir())
    assert len(backups) == 2
    assert backups[0].name != backups[1].name
