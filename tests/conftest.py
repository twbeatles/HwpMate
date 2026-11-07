from __future__ import annotations

import os
import sys
from pathlib import Path

import pytest

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))


@pytest.fixture(scope="session")
def qapp():
    from PySide6.QtWidgets import QApplication

    app = QApplication.instance()
    if app is None:
        app = QApplication([])
    return app


@pytest.fixture(autouse=True)
def _fast_print_output_wait(monkeypatch):
    """PrintToPDFEx 비동기 기록 대기를 테스트에서는 짧게 (실제 대기는 개별 테스트에서 검증)."""
    import hwpmate.services.hwp_print_settings.pdf_export as pdf_export

    monkeypatch.setattr(pdf_export, "PRINT_OUTPUT_WAIT_SECONDS", 0.05)
    monkeypatch.setattr(pdf_export, "PRINT_OUTPUT_POLL_SECONDS", 0.01)
