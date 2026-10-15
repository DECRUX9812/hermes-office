from __future__ import annotations

from pathlib import Path
from typing import Any


def data_dir() -> Path:
    try:
        from plugins.plugin_storage import plugin_data_dir

        return Path(plugin_data_dir("hermes-office"))
    except ImportError:
        pass
    try:
        import os

        home = os.environ.get("HERMES_HOME")
        if home:
            return Path(home).expanduser() / "plugin-data" / "hermes-office"
    except Exception:
        pass
    return Path.home() / ".hermes" / "plugin-data" / "hermes-office"


def emit(event: str, payload: dict[str, Any]) -> None:
    try:
        from hermes_cli.plugin_events import broadcast_plugin_event
    except ImportError:
        return
    try:
        import tui_gateway.server  # noqa: F401
    except Exception:
        return
    broadcast_plugin_event("hermes-office", event, payload)


def conversation(session_id: str) -> dict[str, Any]:
    try:
        from hermes_state import SessionDB

        db = SessionDB()
        messages = db.get_messages(session_id)
    except Exception as exc:
        return {"available": False, "reason": str(exc)}
    convo = []
    for message in messages:
        convo.append(
            {
                "role": message.get("role") or message.get("speaker") or "assistant",
                "text": message.get("text") or message.get("content") or message.get("message") or message.get("body") or "",
                "ts": message.get("ts") or message.get("created_at") or message.get("timestamp") or "",
            }
        )
    return {"available": True, "messages": convo}
