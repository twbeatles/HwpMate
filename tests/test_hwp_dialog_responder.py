from __future__ import annotations

import time

import hwpmate.windows_integration as wi
from hwpmate.windows_integration import hwp_dialog_responder as responder_module
from hwpmate.windows_integration.hwp_dialog_responder import (
    HwpDialogAutoResponder,
    find_hwp_compat_dialogs,
    respond_hwp_compat_dialogs,
)


def _patch_windows(monkeypatch, windows: dict[int, str], *, visible: set[int] | None = None) -> list[set[int]]:
    requested_pids: list[set[int]] = []

    def fake_list(pids, *, security_dialogs_only=False):
        del security_dialogs_only
        requested_pids.append(set(pids))
        return list(windows)

    monkeypatch.setattr(wi, "_list_top_level_hwnds_for_pids", fake_list)
    monkeypatch.setattr(wi, "_window_title", lambda hwnd: windows[hwnd])
    visible_set = set(windows) if visible is None else visible
    monkeypatch.setattr(responder_module, "_is_window_visible", lambda hwnd: hwnd in visible_set)
    return requested_pids


def test_find_compat_dialogs_matches_exact_title_only(monkeypatch) -> None:
    _patch_windows(
        monkeypatch,
        {
            1: "호환 문서",
            2: "빈 문서 1 - 한글",
            3: "호환 문서 저장 옵션",
            4: "보안 승인",
        },
    )

    assert find_hwp_compat_dialogs({100}) == [1]


def test_find_compat_dialogs_requires_owned_pids(monkeypatch) -> None:
    requested = _patch_windows(monkeypatch, {1: "호환 문서"})

    assert find_hwp_compat_dialogs(set()) == []
    assert find_hwp_compat_dialogs(None) == []
    assert requested == []


def test_find_compat_dialogs_skips_hidden_window(monkeypatch) -> None:
    _patch_windows(monkeypatch, {1: "호환 문서", 2: "호환 문서"}, visible={2})

    assert find_hwp_compat_dialogs({100}) == [2]


def test_respond_posts_continue_key_and_honors_skip(monkeypatch) -> None:
    _patch_windows(monkeypatch, {1: "호환 문서", 2: "호환 문서"})
    posted: list[int] = []
    monkeypatch.setattr(responder_module, "_post_continue_key", lambda hwnd: posted.append(hwnd) or True)

    responded = respond_hwp_compat_dialogs({100}, frozenset({2}))

    assert responded == [1]
    assert posted == [1]


def test_auto_responder_counts_responses_with_cooldown() -> None:
    calls: list[tuple[set[int], frozenset[int]]] = []

    def fake_responder(pids, skip):
        calls.append((set(pids), skip))
        return [] if 42 in skip else [42]

    responder = HwpDialogAutoResponder(
        lambda: {7},
        poll_interval=0.02,
        cooldown=60.0,
        responder=fake_responder,
    )
    with responder:
        deadline = time.monotonic() + 1.0
        while len(calls) < 3 and time.monotonic() < deadline:
            time.sleep(0.02)

    assert responder.response_count == 1
    assert calls[0] == ({7}, frozenset())
    assert all(42 in skip for _, skip in calls[1:])


def test_auto_responder_does_nothing_without_pids() -> None:
    calls: list[object] = []
    responder = HwpDialogAutoResponder(
        lambda: set(),
        poll_interval=0.02,
        responder=lambda pids, skip: calls.append(pids) or [1],
    )
    with responder:
        time.sleep(0.1)

    assert calls == []
    assert responder.response_count == 0
