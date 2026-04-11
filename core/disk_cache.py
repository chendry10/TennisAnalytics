from __future__ import annotations

import pickle
from pathlib import Path
from typing import Any


CACHE_ROOT = Path(__file__).resolve().parent.parent / ".cache" / "courtside-analytics"


def _cache_path(cache_key: str) -> Path:
    return CACHE_ROOT / f"{cache_key}.pkl"


def load_cache_entry(cache_key: str) -> Any | None:
    path = _cache_path(cache_key)
    if not path.exists():
        return None

    try:
        with path.open("rb") as handle:
            return pickle.load(handle)
    except Exception:
        return None


def save_cache_entry(cache_key: str, payload: Any) -> None:
    path = _cache_path(cache_key)
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("wb") as handle:
        pickle.dump(payload, handle, protocol=pickle.HIGHEST_PROTOCOL)