"""JSON-backed store for the list of user IDs to register coupons for."""
from __future__ import annotations

import json
import threading

from paths import base_dir

_LOCK = threading.Lock()
_PATH = base_dir() / "ids.json"


def _read() -> list[str]:
    if not _PATH.exists():
        return []
    try:
        data = json.loads(_PATH.read_text(encoding="utf-8"))
    except json.JSONDecodeError:
        return []
    if not isinstance(data, list):
        return []
    return [str(x) for x in data]


def _write(ids: list[str]) -> None:
    _PATH.write_text(
        json.dumps(ids, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


def list_ids() -> list[str]:
    with _LOCK:
        return _read()


def add_ids(new_ids: list[str]) -> tuple[list[str], list[str]]:
    """Add IDs. Returns (added, duplicates)."""
    added: list[str] = []
    dupes: list[str] = []
    with _LOCK:
        current = _read()
        seen = set(current)
        for raw in new_ids:
            v = raw.strip()
            if not v:
                continue
            if v in seen:
                dupes.append(v)
                continue
            current.append(v)
            seen.add(v)
            added.append(v)
        _write(current)
    return added, dupes


def remove_ids(target_ids: list[str]) -> tuple[list[str], list[str]]:
    """Remove IDs. Returns (removed, missing)."""
    removed: list[str] = []
    missing: list[str] = []
    with _LOCK:
        current = _read()
        current_set = set(current)
        for raw in target_ids:
            v = raw.strip()
            if not v:
                continue
            if v in current_set:
                current = [x for x in current if x != v]
                current_set.remove(v)
                removed.append(v)
            else:
                missing.append(v)
        _write(current)
    return removed, missing
