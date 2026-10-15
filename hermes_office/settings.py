"""User settings for office-desk.

They are declared in plugin.yaml ``config_schema``. Hermes builds the settings form from that schema
(Desktop: Settings > Plugins > office-desk) and stores the values under
``plugins.entries.office-desk.settings`` in config.yaml. The plugin has no settings UI of its own.

Inside the agent process the values are read through ``ctx.get_config``; the Desk backend routes
have no plugin context, so they read the same config subtree through ``hermes_cli.config``.
Without Hermes (plain pytest, scripts) every setting falls back to its schema default.
"""

from __future__ import annotations

import logging
import threading
from pathlib import Path
from typing import Any

from .bridge import PLUGIN_ID

logger = logging.getLogger(__name__)

DEFAULTS: dict[str, Any] = {
    "folders": [],
    "default_author": "Hermes",
    "recalc_timeout": 120,
}

_ctx: Any = None
_synced_lock = threading.Lock()
_synced: set[tuple[str, tuple[str, ...]]] = set()


def bind(ctx: Any) -> None:
    """Remember the plugin context so reads go through ``ctx.get_config``."""
    global _ctx
    _ctx = ctx


def _from_config(key: str) -> Any:
    try:
        from hermes_cli.config import load_config_readonly
    except ImportError:
        return None
    try:
        config = load_config_readonly() or {}
    except Exception:
        logger.debug("office-desk: could not read config.yaml", exc_info=True)
        return None
    entry = (((config.get("plugins") or {}).get("entries") or {}).get(PLUGIN_ID)) or {}
    settings = entry.get("settings") if isinstance(entry, dict) else None
    return settings.get(key) if isinstance(settings, dict) else None


def get(key: str) -> Any:
    """The configured value of ``key``, or its default when unset or of the wrong type."""
    default = DEFAULTS[key]
    value: Any = None
    if _ctx is not None:
        try:
            value = _ctx.get_config(key, None)
        except Exception:
            value = None
    else:
        value = _from_config(key)
    if value is None:
        return default
    if isinstance(default, bool) or isinstance(value, bool):
        return value if isinstance(value, bool) and isinstance(default, bool) else default
    if isinstance(default, int):
        return value if isinstance(value, int) and value > 0 else default
    if not isinstance(value, type(default)):
        logger.warning("office-desk setting %s should be %s; using the default", key, type(default).__name__)
        return default
    return value


def default_author() -> str:
    author = get("default_author")
    return author.strip() or DEFAULTS["default_author"]


def recalc_timeout() -> float:
    return float(get("recalc_timeout"))


def folders() -> list[Path]:
    """Configured Desk folders that exist (others are skipped with a warning)."""
    out: list[Path] = []
    for raw in get("folders"):
        if not isinstance(raw, str) or not raw.strip():
            continue
        path = Path(raw).expanduser()
        if not path.is_absolute() or not path.is_dir():
            logger.warning("office-desk: Desk folder %r is not an existing absolute folder; skipped", raw)
            continue
        out.append(path.resolve())
    return out


def sync_folders(desk: Any) -> None:
    """Put the configured folders on ``desk`` (once per desk and folder list in this process)."""
    wanted = folders()
    key = (str(desk.db_path), tuple(str(p) for p in wanted))
    with _synced_lock:
        if key in _synced:
            return
        _synced.add(key)
    if not wanted:
        return
    present = {folder["path"] for folder in desk.list_folders()}
    for path in wanted:
        if str(path) not in present:
            desk.add_folder(path)


def _reset_for_tests() -> None:
    global _ctx
    _ctx = None
    with _synced_lock:
        _synced.clear()
