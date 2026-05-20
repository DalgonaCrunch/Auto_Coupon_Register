"""Resolve runtime data paths consistently for both `python coupon_bot.py`
and PyInstaller-frozen executables.

When frozen, `sys.executable` points at the .exe and `_MEIPASS` is a temp
extraction dir; data files belong NEXT TO the exe, not inside _MEIPASS.
"""
from __future__ import annotations

import sys
from pathlib import Path


def base_dir() -> Path:
    """Directory where ids.json / bot_state.json / .env live."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent
