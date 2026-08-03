from __future__ import annotations

from hwpmate.constants import HWP_FOREGROUND_POLL_MS, HWP_FOREGROUND_POLL_MS_RELAXED
from hwpmate.services.hwp_security_session import HwpSecuritySession


def test_target_pids_prefers_owned() -> None:
    session = HwpSecuritySession(owned_pids={10, 20})
    assert session.target_pids() == {10, 20}


def test_target_pids_none_when_empty() -> None:
    session = HwpSecuritySession()
    assert session.target_pids() is None


def test_should_auto_accept_disabled_when_module_registered() -> None:
    session = HwpSecuritySession(
        auto_accept_enabled=True,
        security_module_registered=True,
        engine_status_received=True,
    )
    assert session.should_auto_accept() is False


def test_should_auto_accept_blocked_before_engine_status() -> None:
    session = HwpSecuritySession(
        auto_accept_enabled=True,
        security_module_registered=None,
        engine_status_received=False,
    )
    assert session.should_auto_accept() is False


def test_should_auto_accept_blocked_when_registered_unknown() -> None:
    session = HwpSecuritySession(
        auto_accept_enabled=True,
        security_module_registered=None,
        engine_status_received=True,
    )
    assert session.should_auto_accept() is False


def test_should_auto_accept_respects_user_toggle() -> None:
    session = HwpSecuritySession(
        auto_accept_enabled=False,
        engine_status_received=True,
        security_module_registered=False,
    )
    assert session.should_auto_accept() is False


def test_should_auto_accept_cooldown_and_note() -> None:
    session = HwpSecuritySession(
        auto_accept_enabled=True,
        security_module_registered=False,
        engine_status_received=True,
        cooldown_seconds=60.0,
    )
    assert session.should_auto_accept() is True
    session.note_auto_accept()
    assert session.auto_accept_clicks == 1
    assert session.should_auto_accept() is False


def test_poll_interval_relaxes_when_module_ok() -> None:
    session = HwpSecuritySession(security_module_registered=True)
    assert session.poll_interval_ms() == HWP_FOREGROUND_POLL_MS_RELAXED
    session.security_module_registered = False
    assert session.poll_interval_ms() == HWP_FOREGROUND_POLL_MS


def test_apply_engine_status() -> None:
    session = HwpSecuritySession()
    assert session.engine_status_received is False
    session.apply_engine_status(
        {
            "security_module_registered": True,
            "owned_pids": [1, 2, "3"],
            "snapshot_unreliable": True,
        }
    )
    assert session.security_module_registered is True
    assert session.owned_pids == {1, 2, 3}
    assert session.snapshot_unreliable is True
    assert session.engine_status_received is True
