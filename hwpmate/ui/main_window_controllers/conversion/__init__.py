"""conversion controller package.

호환: `from hwpmate.ui.main_window_controllers.conversion import ConversionController`
"""

from __future__ import annotations

from .controller import ConversionController

from ....windows_integration import (  # noqa: F401 -- monkeypatched test seam (api.* in controller, tests/test_main_window_controllers.py)
    bring_hwp_windows_to_foreground,
    hide_hwp_main_windows,
    try_accept_hwp_security_dialog,
)


__all__ = ["ConversionController"]
