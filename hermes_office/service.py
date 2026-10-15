"""Desk-aware operations shared by the agent tools and the Desk backend routes.

Every write goes: policy checks -> backup -> library edit (atomic save) -> receipt -> events.
A failure before or inside the library edit leaves the file byte-identical.
"""

from __future__ import annotations

import threading
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Callable
from uuid import uuid4

from . import bridge
from .desk import Desk
from .docx_ops import resolve_revisions_docx
from .errors import NotFoundError, OutsideDeskError
from .guard import atomic_write, backup, kind_of, require_kind, resolve_path, sha256

TITLE_LIMIT = 60

# Turn state lives in the agent process: pre_llm_call/post_llm_call hooks mark which sessions are mid-turn,
# so a finished tool keeps its workstream "running" until the agent's turn actually ends.
_turn_lock = threading.Lock()
_active_turns: dict[str, str] = {}


def turn_started(session_id: str, user_message: str) -> None:
    with _turn_lock:
        _active_turns[session_id] = user_message or ""


def turn_ended(session_id: str) -> str | None:
    with _turn_lock:
        return _active_turns.pop(session_id, None)


def _turn_message(session_id: str | None) -> str | None:
    if not session_id:
        return None
    with _turn_lock:
        return _active_turns.get(session_id)


def desk() -> Desk:
    return Desk(bridge.data_dir())


def notify(ws_id: int, status: str | None = None) -> None:
    bridge.emit("workstream.updated", {"workstream_id": ws_id, "status": status})


def set_status(d: Desk, ws_id: int, status: str, *, note: str | None = None, needs_you: bool | None = None) -> dict[str, Any]:
    fields: dict[str, Any] = {"status": status}
    if note is not None:
        fields["note"] = note
    if needs_you is not None:
        fields["needs_you"] = 1 if needs_you else 0
    ws = d.update_workstream(ws_id, **fields)
    notify(ws_id, status)
    return ws


def _title_from(message: str | None, fallback: str) -> str:
    text = " ".join((message or "").split())
    if not text:
        return fallback
    return text if len(text) <= TITLE_LIMIT else text[: TITLE_LIMIT - 1].rstrip() + "…"


def resolve_workstream(d: Desk, *, path: Path, workstream_id: Any, session_id: str | None, fallback_title: str) -> dict[str, Any]:
    """Explicit id -> the calling session's workstream -> a new one on the folder that holds `path`."""
    folder = d.folder_for_path(path)
    if folder is None:
        raise OutsideDeskError(f"{path} is not inside a Desk folder; nothing was changed")
    if workstream_id is not None:
        ws = d.get_workstream(int(workstream_id))
        if not ws:
            raise NotFoundError(f"workstream {workstream_id} not found; nothing was changed")
        if session_id and not ws.get("session_id"):
            ws = d.update_workstream(ws["id"], session_id=session_id)
        return ws
    if session_id:
        ws = d.workstream_for_session(session_id)
        if ws:
            return ws
    message = _turn_message(session_id)
    return d.create_workstream(folder["id"], _title_from(message, fallback_title), request=message, session_id=session_id, status="running")


@dataclass
class WriteOp:
    tool: str
    path: str
    kind: str
    run: Callable[[], dict[str, Any]]
    summary: str
    output_path: str | None = None
    needs_you: str | None = None  # note shown in "Needs you" when the result awaits review


def run_write(op: WriteOp, *, workstream_id: Any = None, session_id: str | None = None, d: Desk | None = None) -> dict[str, Any]:
    d = d or desk()
    source = resolve_path(op.path)
    target = resolve_path(op.output_path) if op.output_path else source
    ws = None
    op_id = uuid4().hex
    try:
        require_kind(source, op.kind)
        require_kind(target, op.kind)
        if not source.is_file():
            raise FileNotFoundError(f"{source} does not exist")
        ws = resolve_workstream(d, path=target, workstream_id=workstream_id, session_id=session_id, fallback_title=f"{op.summary} · {target.name}")
        d.add_event(ws["id"], "tool_started", tool=op.tool, path=str(target), summary=op.summary, data={"op_id": op_id})
        if not ws.get("needs_you"):
            set_status(d, ws["id"], "running")
        existed = target.exists()
        before = sha256(target) if existed else None
        backup_path = backup(d.data_dir, ws["id"], target) if existed else None
        result = op.run()
    except Exception as exc:
        if ws is not None:
            d.add_event(ws["id"], "refused", tool=op.tool, path=str(target), summary=f"{type(exc).__name__}: {exc}", data={"op_id": op_id})
            notify(ws["id"])
        raise
    after = sha256(target)
    receipt = None
    if after != before:
        receipt = d.add_receipt(ws["id"], op.tool, target, op.summary, before, after, backup_path, result)
        d.track_file(ws["id"], target, "changed" if existed else "created", receipt_id=receipt["id"], kind=op.kind)
    d.add_event(ws["id"], "tool_finished", tool=op.tool, path=str(target), summary=op.summary if receipt else f"{op.summary} (no change)", data={"op_id": op_id, "receipt_id": receipt["id"] if receipt else None})
    if op.needs_you and receipt:
        set_status(d, ws["id"], "needs_you", note=op.needs_you, needs_you=True)
    elif not ws.get("needs_you"):
        set_status(d, ws["id"], "running" if _turn_message(session_id) is not None else "done")
    return {"workstream_id": ws["id"], "receipt_id": receipt["id"] if receipt else None, "changed": receipt is not None, "result": result}


def note_read(d: Desk, *, tool: str, path: str, session_id: str | None, workstream_id: Any = None) -> None:
    """Mark a file as being worked on by the calling session's workstream (reads never create workstreams)."""
    ws = d.get_workstream(int(workstream_id)) if workstream_id is not None else (d.workstream_for_session(session_id) if session_id else None)
    if not ws:
        return
    resolved = resolve_path(path)
    try:
        kind = kind_of(resolved)
    except Exception:
        return
    known = {f["path"]: f for f in d.files(ws["id"])}
    if str(resolved) not in known and d.folder_for_path(resolved) is not None:
        d.track_file(ws["id"], resolved, "working", kind=kind)
    op_id = uuid4().hex
    d.add_event(ws["id"], "tool_finished", tool=tool, path=str(resolved), summary=f"Read {resolved.name}", data={"op_id": op_id})
    notify(ws["id"])


def revert_receipt(receipt_id: int, *, d: Desk | None = None) -> dict[str, Any]:
    d = d or desk()
    receipt = d.get_receipt(int(receipt_id))
    if receipt is None:
        raise NotFoundError(f"receipt {receipt_id} not found")
    if receipt.get("reverted_at"):
        raise NotFoundError(f"receipt {receipt_id} was already reverted at {receipt['reverted_at']}")
    target = Path(receipt["path"])
    source_backup = receipt.get("backup_path")
    if source_backup is None:
        if not target.exists():
            raise NotFoundError(f"{target} no longer exists")
        # The receipt created the file: reverting removes it, after keeping a copy.
        keep = backup(d.data_dir, receipt["workstream_id"], target)
        before = sha256(target)
        target.unlink()
        new = d.add_receipt(receipt["workstream_id"], "office_receipt", target, f"Revert receipt {receipt_id} (remove created file)", before, None, keep, {"reverted": receipt_id})
    else:
        if not Path(source_backup).exists():
            raise NotFoundError(f"backup for receipt {receipt_id} is missing: {source_backup}")
        keep = backup(d.data_dir, receipt["workstream_id"], target) if target.exists() else None
        before = sha256(target) if target.exists() else None
        data = Path(source_backup).read_bytes()
        atomic_write(target, lambda tmp: tmp.write_bytes(data))
        new = d.add_receipt(receipt["workstream_id"], "office_receipt", target, f"Revert receipt {receipt_id}", before, sha256(target), keep, {"reverted": receipt_id})
        d.track_file(receipt["workstream_id"], target, "changed", receipt_id=new["id"], kind=kind_of(target))
    d.mark_reverted(int(receipt_id))
    d.add_event(receipt["workstream_id"], "note", tool="office_receipt", path=str(target), summary=f"Reverted receipt {receipt_id}", data={"op_id": uuid4().hex})
    notify(receipt["workstream_id"])
    return new


def resolve_redline(path: str, *, action: str, workstream_id: Any = None, author: str | None = None, scope: str = "all", index: int | None = None, session_id: str | None = None, d: Desk | None = None) -> dict[str, Any]:
    d = d or desk()
    verb = "Accept" if action == "accept" else "Reject"
    op = WriteOp(
        tool="docx_redline_accept_reject",
        path=path,
        kind="docx",
        summary=f"{verb} tracked changes" + (f" by {author}" if scope == "author" and author else ""),
        run=lambda: resolve_revisions_docx(path, action=action, scope=scope, author=author, index=index),
    )
    out = run_write(op, workstream_id=workstream_id, session_id=session_id, d=d)
    if out["result"].get("remaining") == 0:
        ws = d.get_workstream(out["workstream_id"])
        if ws and ws.get("needs_you"):
            set_status(d, ws["id"], "done", note=f"{verb}ed {out['result']['resolved']} change(s)", needs_you=False)
    return out
