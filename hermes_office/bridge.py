"""The only module that touches Hermes. Every touch point is probed so the core runs without Hermes."""

from __future__ import annotations

import logging
import os
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

PLUGIN_ID = "hermes-office"

logger = logging.getLogger(__name__)


def data_dir() -> Path:
    try:
        from plugins.plugin_storage import plugin_data_dir
    except ImportError:
        home = Path(os.environ.get("HERMES_HOME") or Path.home() / ".hermes").expanduser()
        return home / "plugin-data" / PLUGIN_ID
    return Path(plugin_data_dir(PLUGIN_ID))


def emit(event: str, payload: dict[str, Any]) -> None:
    """Push a live update to Desktop clients; a no-op outside `hermes serve`."""
    try:
        from hermes_cli.plugin_events import broadcast_plugin_event
    except ImportError:
        return
    try:
        broadcast_plugin_event(PLUGIN_ID, event, payload)
    except Exception:  # live push is best-effort; the Desk also polls
        logger.debug("hermes-office event push failed", exc_info=True)


def _message_text(content: Any) -> str:
    if isinstance(content, str):
        return content
    if isinstance(content, list):
        return "\n".join(part.get("text", "") for part in content if isinstance(part, dict) and part.get("type") in (None, "text"))
    return "" if content is None else str(content)


def _iso(ts: Any) -> str | None:
    if isinstance(ts, (int, float)):
        return datetime.fromtimestamp(ts, timezone.utc).isoformat(timespec="seconds")
    return ts if isinstance(ts, str) else None


def conversation(session_id: str, *, limit: int = 200) -> dict[str, Any]:
    try:
        from hermes_state import SessionDB
    except ImportError:
        return {"available": False, "reason": "Hermes session store is not importable in this process", "messages": []}
    try:
        db = SessionDB(read_only=True)
        rows = db.get_messages(session_id, include_ancestors=True)
    except Exception as exc:
        return {"available": False, "reason": f"could not read session {session_id}: {exc}", "messages": []}
    messages = []
    for row in rows[-limit:]:
        role = row.get("role")
        if role == "tool":
            messages.append({"role": "tool", "text": row.get("tool_name") or "tool result", "ts": _iso(row.get("timestamp"))})
            continue
        text = _message_text(row.get("content"))
        if role == "assistant" and not text.strip() and row.get("tool_calls"):
            names = [call.get("function", {}).get("name", "tool") for call in row["tool_calls"] if isinstance(call, dict)]
            text = "Calling " + ", ".join(names)
            role = "tool"
        if text.strip():
            messages.append({"role": role, "text": text, "ts": _iso(row.get("timestamp"))})
    if not messages:
        return {"available": False, "reason": f"session {session_id} has no messages yet", "messages": []}
    return {"available": True, "messages": messages}
