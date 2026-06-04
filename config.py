"""Local persistence for AI settings — API keys, model, prompts, batch size.

Stored as JSON in ``~/.pdf2excel.json`` so it survives across runs and across
working directories (the GUI is a desktop app). **The file can contain API keys**
— it is gitignored and written with ``0600`` permissions; never commit it.

API keys resolve in order: saved config → environment variable
(``ANTHROPIC_API_KEY`` / ``OPENAI_API_KEY`` / ``GOOGLE_API_KEY``).

Public API:
    load() -> dict
    save(data) -> None
    get_api_key(provider, data=None) -> str
    set_api_key(provider, key, data=None) -> dict
"""

from __future__ import annotations

import json
import os
from typing import Any, Dict, Optional

CONFIG_FILENAME = ".pdf2excel.json"

# provider name -> environment variable that may hold its key (fallback).
_ENV_KEYS = {
    "claude": "ANTHROPIC_API_KEY",
    "openai": "OPENAI_API_KEY",
    "gemini": "GOOGLE_API_KEY",
}


def config_path() -> str:
    """Absolute path to the settings file (``~/.pdf2excel.json``)."""
    return os.path.join(os.path.expanduser("~"), CONFIG_FILENAME)


def load() -> Dict[str, Any]:
    """Load the settings dict (``{}`` if the file is missing or invalid)."""
    try:
        with open(config_path(), encoding="utf-8") as fh:
            data = json.load(fh)
        return data if isinstance(data, dict) else {}
    except (OSError, ValueError):
        return {}


def save(data: Dict[str, Any]) -> None:
    """Persist ``data`` as JSON, restricting permissions (it may hold keys)."""
    path = config_path()
    try:
        with open(path, "w", encoding="utf-8") as fh:
            json.dump(data, fh, indent=2)
        try:
            os.chmod(path, 0o600)
        except OSError:
            pass  # best-effort on platforms without POSIX permissions
    except OSError as exc:
        raise RuntimeError(f"could not save settings to {path}: {exc}") from exc


def get_api_key(provider: str, data: Optional[Dict[str, Any]] = None) -> str:
    """Return the API key for ``provider`` (saved config first, then env var)."""
    data = load() if data is None else data
    saved = (data.get("api_keys") or {}).get(provider, "")
    if saved:
        return saved
    return os.environ.get(_ENV_KEYS.get(provider, ""), "")


def set_api_key(provider: str, key: str,
                data: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
    """Set/clear ``provider``'s key in ``data`` (loads + saves if not provided)."""
    own = data is None
    data = load() if own else data
    keys = data.setdefault("api_keys", {})
    if key:
        keys[provider] = key
    else:
        keys.pop(provider, None)
    if own:
        save(data)
    return data
