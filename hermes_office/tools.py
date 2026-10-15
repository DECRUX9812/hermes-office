"""Agent tool handlers: `handler(args, **kwargs) -> JSON string`. They never raise."""

from __future__ import annotations

import functools
import json
from pathlib import Path
from typing import Any, Callable

from . import service
from .docx_ops import comment_docx, redline_docx, tracked_edit_docx
from .errors import NotFoundError, OutsideDeskError
from .guard import refusal_payload, resolve_path
from .inspect_ops import inspect_file
from .pptx_ops import compose_pptx_decks, diff_pptx, edit_pptx_shape, inspect_pptx_formats
from .search_ops import search_path_or_folder
from .service import WriteOp
from .xlsx_ops import diff_xlsx, edit_xlsx, recalc_xlsx

Handler = Callable[..., str]


def _json(payload: Any) -> str:
    return json.dumps(payload, ensure_ascii=False, default=str)


def tool(fn: Callable[..., dict[str, Any]]) -> Handler:
    @functools.wraps(fn)
    def handler(args: dict[str, Any], **kwargs: Any) -> str:
        try:
            return _json({"ok": True, **fn(args or {}, kwargs.get("session_id"))})
        except Exception as exc:
            return _json(refusal_payload(exc))

    return handler


def _write(op: WriteOp, args: dict[str, Any], session_id: str | None) -> dict[str, Any]:
    out = service.run_write(op, workstream_id=args.get("workstream_id"), session_id=session_id)
    result = out.pop("result")
    return {**out, **result}


def _current_ws(d, args: dict[str, Any], session_id: str | None) -> dict[str, Any]:
    if args.get("workstream_id") is not None:
        ws = d.get_workstream(int(args["workstream_id"]))
    else:
        ws = d.workstream_for_session(session_id) if session_id else None
    if not ws:
        raise NotFoundError("no workstream: pass workstream_id, or edit a Desk file first (that opens one for this chat)")
    return ws


@tool
def office_desk(args: dict[str, Any], session_id: str | None) -> dict[str, Any]:
    d = service.desk()
    action = args["action"]
    if action == "list":
        return {"folders": d.list_folders(), "workstreams": d.list_workstreams(limit=20)}
    if action == "add_folder":
        path = resolve_path(args["path"])
        if not path.is_dir():
            raise FileNotFoundError(f"{path} is not a folder")
        return {"folder": d.add_folder(path)}
    if action == "remove_folder":
        if d.get_folder(int(args["folder_id"])) is None:
            raise NotFoundError(f"folder {args['folder_id']} not found")
        d.remove_folder(int(args["folder_id"]))
        return {"removed": int(args["folder_id"])}
    if action == "status":
        ws = _current_ws(d, args, session_id)
        return {"workstream": ws, "now": d.now(ws["id"]), "files": d.files(ws["id"])}
    if action == "needs_you":
        ws = _current_ws(d, args, session_id)
        return {"workstream": service.set_status(d, ws["id"], "needs_you", note=args.get("note") or "Please review", needs_you=True)}
    if action == "done":
        ws = _current_ws(d, args, session_id)
        return {"workstream": service.set_status(d, ws["id"], "done", note=args.get("note"), needs_you=False)}
    if action == "attach":
        ws = _current_ws(d, {"workstream_id": args.get("workstream_id")}, None)
        fields = {"session_id": session_id} if session_id else {}
        if args.get("title"):
            fields["title"] = args["title"]
        ws = d.update_workstream(ws["id"], **fields)
        service.notify(ws["id"])
        return {"workstream": ws}
    raise ValueError(f"unknown action {action!r}")


@tool
def office_inspect(args: dict[str, Any], session_id: str | None) -> dict[str, Any]:
    result = inspect_file(args["path"], detail=args.get("detail") or "summary")
    service.note_read(service.desk(), tool="office_inspect", path=args["path"], session_id=session_id, workstream_id=args.get("workstream_id"))
    return result


@tool
def office_search(args: dict[str, Any], session_id: str | None) -> dict[str, Any]:
    results = search_path_or_folder(args.get("path"), args.get("folder"), query=args["query"], match=args.get("match") or "exact", regex=bool(args.get("regex")))
    return {"results": results, "files_with_matches": sum(1 for r in results if r.get("matches"))}


@tool
def docx_tracked_edit(args: dict[str, Any], session_id: str | None) -> dict[str, Any]:
    path = args["path"]
    name = Path(path).name
    op = WriteOp(
        tool="docx_tracked_edit",
        path=path,
        kind="docx",
        summary=f"Tracked change: {args['find'][:40]!r} → {args['replace'][:40]!r}",
        needs_you=f"Review tracked changes in {name}",
        run=lambda: tracked_edit_docx(path, find=args["find"], replace=args["replace"], nth=args.get("nth"), match=args.get("match") or "exact", author=args.get("author")),
    )
    return _write(op, args, session_id)


@tool
def docx_comment(args: dict[str, Any], session_id: str | None) -> dict[str, Any]:
    path = args["path"]
    if args.get("resolve") is not None:
        summary = f"Resolve comment {args['resolve']}"
    elif args.get("reply_to") is not None:
        summary = f"Reply to comment {args['reply_to']}"
    else:
        summary = f"Comment on {str(args.get('find', ''))[:40]!r}"
    op = WriteOp(
        tool="docx_comment",
        path=path,
        kind="docx",
        summary=summary,
        run=lambda: comment_docx(path, text=args.get("text") or "", find=args.get("find"), author=args.get("author"), nth=args.get("nth"), reply_to=args.get("reply_to"), resolve_id=args.get("resolve")),
    )
    return _write(op, args, session_id)


@tool
def docx_redline_accept_reject(args: dict[str, Any], session_id: str | None) -> dict[str, Any]:
    path = args["path"]
    action = args["action"]
    if action == "compare":
        original = args.get("compare_with")
        if not original:
            raise ValueError("compare needs compare_with (the original document)")
        revised = resolve_path(path)
        output = args.get("output_path") or str(revised.with_name(f"{revised.stem} (redline).docx"))
        op = WriteOp(
            tool="docx_redline_accept_reject",
            path=str(revised),
            output_path=output,
            kind="docx",
            summary=f"Redline {Path(original).name} → {revised.name}",
            needs_you=f"Review the redline {Path(output).name}",
            run=lambda: redline_docx(original, revised, output, author=args.get("author")),
        )
        return _write(op, args, session_id)
    if action not in ("accept", "reject"):
        raise ValueError("action must be accept, reject or compare")
    out = service.resolve_redline(path, action=action, workstream_id=args.get("workstream_id"), author=args.get("author"), scope=args.get("scope") or "all", index=args.get("index"), session_id=session_id)
    result = out.pop("result")
    return {**out, **result}


@tool
def pptx_inspect_formats(args: dict[str, Any], session_id: str | None) -> dict[str, Any]:
    return inspect_pptx_formats(args["path"], slide=int(args["slide"]), shape=args["shape"])


@tool
def pptx_edit_shape(args: dict[str, Any], session_id: str | None) -> dict[str, Any]:
    path = args["path"]
    ops = list(args["ops"])
    names = ", ".join(str(o.get("op")) for o in ops)
    target = f" {args['shape']!r}" if args.get("shape") else ""
    op = WriteOp(
        tool="pptx_edit_shape",
        path=path,
        kind="pptx",
        summary=f"Slide {int(args['slide']) + 1}{target}: {names}",
        run=lambda: edit_pptx_shape(path, slide=int(args["slide"]), shape=args.get("shape"), ops=ops),
    )
    return _write(op, args, session_id)


@tool
def pptx_compose_decks(args: dict[str, Any], session_id: str | None) -> dict[str, Any]:
    path = args["path"]
    output = args.get("output_path") or path
    sources = list(args["sources"])
    for source in sources:
        if not resolve_path(source["path"]).is_file():
            raise FileNotFoundError(f"source deck {source['path']} does not exist")
    op = WriteOp(
        tool="pptx_compose_decks",
        path=path,
        output_path=output,
        kind="pptx",
        summary=f"Import slides from {', '.join(Path(s['path']).name for s in sources)} ({args['mode']})",
        run=lambda: compose_pptx_decks(path, sources, mode=args["mode"], output_path=output),
    )
    return _write(op, args, session_id)


@tool
def pptx_diff(args: dict[str, Any], session_id: str | None) -> dict[str, Any]:
    return {"diff": diff_pptx(args["a"], args["b"], detail=args.get("detail") or "text")}


@tool
def xlsx_edit(args: dict[str, Any], session_id: str | None) -> dict[str, Any]:
    path = args["path"]
    ops = list(args["ops"])
    refs = [o.get("ref") or o.get("op") for o in ops][:4]
    op = WriteOp(
        tool="xlsx_edit",
        path=path,
        kind="xlsx",
        summary=f"Edit {', '.join(str(r) for r in refs)}" + ("…" if len(ops) > 4 else ""),
        run=lambda: edit_xlsx(path, ops=ops),
    )
    return _write(op, args, session_id)


@tool
def xlsx_recalc(args: dict[str, Any], session_id: str | None) -> dict[str, Any]:
    path = args["path"]
    op = WriteOp(tool="xlsx_recalc", path=path, kind="xlsx", summary="Recalculate formulas", run=lambda: recalc_xlsx(path, read=args.get("read") or []))
    return _write(op, args, session_id)


@tool
def xlsx_diff(args: dict[str, Any], session_id: str | None) -> dict[str, Any]:
    return {"diff": diff_xlsx(args["a"], args["b"])}


@tool
def office_receipt(args: dict[str, Any], session_id: str | None) -> dict[str, Any]:
    d = service.desk()
    action = args["action"]
    if action == "list":
        ws_id = args.get("workstream_id")
        if ws_id is None and session_id:
            ws = d.workstream_for_session(session_id)
            ws_id = ws["id"] if ws else None
        receipts = d.list_receipts(int(ws_id)) if ws_id is not None else d.list_receipts()[:50]
        return {"receipts": [{k: v for k, v in r.items() if k != "library_report"} for r in receipts]}
    if action == "show":
        receipt = d.get_receipt(int(args["receipt_id"]))
        if receipt is None:
            raise NotFoundError(f"receipt {args['receipt_id']} not found")
        return {"receipt": receipt}
    if action == "revert":
        receipt = d.get_receipt(int(args["receipt_id"]))
        if receipt is None:
            raise NotFoundError(f"receipt {args['receipt_id']} not found")
        if d.folder_for_path(receipt["path"]) is None:
            raise OutsideDeskError(f"{receipt['path']} is no longer inside a Desk folder; nothing was changed")
        return {"receipt": service.revert_receipt(int(args["receipt_id"]), d=d)}
    raise ValueError(f"unknown action {action!r}")


HANDLERS: dict[str, Handler] = {
    "office_desk": office_desk,
    "office_inspect": office_inspect,
    "office_search": office_search,
    "docx_tracked_edit": docx_tracked_edit,
    "docx_comment": docx_comment,
    "docx_redline_accept_reject": docx_redline_accept_reject,
    "pptx_inspect_formats": pptx_inspect_formats,
    "pptx_edit_shape": pptx_edit_shape,
    "pptx_compose_decks": pptx_compose_decks,
    "pptx_diff": pptx_diff,
    "xlsx_edit": xlsx_edit,
    "xlsx_recalc": xlsx_recalc,
    "xlsx_diff": xlsx_diff,
    "office_receipt": office_receipt,
}
