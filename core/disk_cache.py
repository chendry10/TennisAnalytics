from __future__ import annotations

import io
import logging
import os
import re
import time
from pathlib import Path
from typing import Any

import pandas as pd


logger = logging.getLogger(__name__)

CACHE_ENV_VAR = "COURTSIDE_ANALYTICS_CACHE_DIR"
MAX_AGE_DAYS_ENV_VAR = "COURTSIDE_ANALYTICS_CACHE_MAX_AGE_DAYS"
DEFAULT_MAX_AGE_DAYS = 30
_CACHE_KEY_PATTERN = re.compile(r"^[a-f0-9]{64}$")


def _default_cache_root() -> Path:
    override = os.environ.get(CACHE_ENV_VAR)
    if override:
        return Path(override).expanduser()

    if os.name == "nt":
        base = os.environ.get("LOCALAPPDATA") or Path.home() / "AppData" / "Local"
        return Path(base) / "CourtSideAnalytics" / "Cache"

    base = os.environ.get("XDG_CACHE_HOME")
    if base:
        return Path(base) / "courtside-analytics"
    return Path.home() / ".cache" / "courtside-analytics"


CACHE_ROOT = _default_cache_root()


def _cache_path(cache_key: str) -> Path:
    if not _CACHE_KEY_PATTERN.fullmatch(cache_key):
        raise ValueError("Invalid cache key")
    return CACHE_ROOT / f"{cache_key}.json"


def _max_age_seconds() -> int:
    try:
        days = int(os.environ.get(MAX_AGE_DAYS_ENV_VAR, DEFAULT_MAX_AGE_DAYS))
    except ValueError:
        days = DEFAULT_MAX_AGE_DAYS
    return max(days, 1) * 24 * 60 * 60


def _prune_old_entries() -> None:
    cutoff = time.time() - _max_age_seconds()
    try:
        if not CACHE_ROOT.exists():
            return
        for path in CACHE_ROOT.glob("*.json"):
            try:
                if path.stat().st_mtime < cutoff:
                    path.unlink()
            except OSError:
                continue
    except OSError as exc:
        logger.debug("Cache cleanup skipped: %s", exc)


def load_cache_entry(cache_key: str) -> Any | None:
    try:
        path = _cache_path(cache_key)
    except ValueError:
        return None

    if not path.exists():
        return None

    try:
        raw = path.read_text(encoding="utf-8")
        return pd.read_json(io.StringIO(raw), orient="table")
    except Exception as exc:
        logger.debug("Ignoring unreadable cache entry %s: %s", path, exc)
        return None


def save_cache_entry(cache_key: str, payload: Any) -> None:
    if not isinstance(payload, pd.DataFrame):
        return

    try:
        path = _cache_path(cache_key)
        path.parent.mkdir(parents=True, exist_ok=True)
        tmp_path = path.with_suffix(".json.tmp")
        tmp_path.write_text(payload.to_json(orient="table"), encoding="utf-8")
        tmp_path.replace(path)
        _prune_old_entries()
    except Exception as exc:
        logger.debug("Cache write skipped: %s", exc)
