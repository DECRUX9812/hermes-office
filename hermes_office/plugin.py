from __future__ import annotations

from pathlib import Path
from typing import Any

from . import service, settings
from .schemas import ALL_SCHEMAS
from .tools import HANDLERS

TOOLSET = "office"
SKILL_DIR = Path(__file__).parent / "skills" / "office-review"


def engines_ready() -> bool:
    """True when the Paper forks (not stock python-docx/pptx/openpyxl) are the importable packages."""
    try:
        import docx
        import openpyxl
        import pptx
    except ImportError:
        return False
    return all(getattr(module, "__paper_version__", None) for module in (docx, pptx, openpyxl))


def _on_turn_start(session_id: str = "", user_message: Any = None, **_: Any) -> None:
    if session_id:
        service.turn_started(session_id, user_message if isinstance(user_message, str) else "")
        d = service.desk()
        ws = d.workstream_for_session(session_id)
        if ws and not ws.get("needs_you") and ws.get("status") != "running":
            service.set_status(d, ws["id"], "running")


def _on_turn_end(session_id: str = "", **_: Any) -> None:
    if not session_id or service.turn_ended(session_id) is None:
        return
    d = service.desk()
    ws = d.workstream_for_session(session_id)
    if not ws or ws.get("needs_you"):
        return
    last = ws.get("last_event") or {}
    if last.get("kind") == "refused":
        service.set_status(d, ws["id"], "needs_you", note=f"An edit was refused: {last.get('summary', '')}"[:300], needs_you=True)
    else:
        service.set_status(d, ws["id"], "done")


def register(ctx) -> None:
    settings.bind(ctx)
    for name, handler in HANDLERS.items():
        ctx.register_tool(name=name, toolset=TOOLSET, schema=ALL_SCHEMAS[name], handler=handler, check_fn=engines_ready, emoji="🗂️")
    ctx.register_skill("office-review", SKILL_DIR / "SKILL.md", description="Review and edit Word, PowerPoint and Excel files on the Desk with tracked changes and receipts.")
    ctx.register_hook("pre_llm_call", _on_turn_start)
    ctx.register_hook("post_llm_call", _on_turn_end)
