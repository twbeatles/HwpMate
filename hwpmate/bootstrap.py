from __future__ import annotations

import os

os.environ["QT_ENABLE_HIGHDPI_SCALING"] = "1"
os.environ["QT_AUTO_SCREEN_SCALE_FACTOR"] = "1"

from .app import main

__all__ = ["main"]
