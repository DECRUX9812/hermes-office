"""Desk backend routes, mounted by Hermes at /api/plugins/office-desk/ (dashboard and `hermes serve`)."""

from __future__ import annotations

import importlib.util
import sys
from pathlib import Path
from typing import Any

from fastapi import APIRouter, HTTPException
from pydantic import BaseModel

_ROOT = Path(__file__).resolve().parents[1]


def _load_core() -> None:
    # Hermes loads this file by path, so the plugin's own package may not be importable by name.
    try:
        import hermes_office  # noqa: F401
        return
    except ImportError:
        pass
    spec = importlib.util.spec_from_file_location("hermes_office", _ROOT / "hermes_office" / "__init__.py", submodule_search_locations=[str(_ROOT / "hermes_office")])
    module = importlib.util.module_from_spec(spec)
    sys.modules["hermes_office"] = module
    spec.loader.exec_module(module)


_load_core()

from hermes_office import __version__, bridge, service  # noqa: E402
from hermes_office.desk import Desk  # noqa: E402
from hermes_office.diff_ops import diff_file  # noqa: E402
from hermes_office.errors import OfficeError  # noqa: E402
from hermes_office.guard import OFFICE_SUFFIXES, kind_of, refusal_payload, resolve_path  # noqa: E402
from hermes_office.preview import preview_file  # noqa: E402
from hermes_office.search_ops import office_files  # noqa: E402
from hermes_office.xlsx_ops import soffice_path  # noqa: E402

router = APIRouter()

STATUSES = {"running", "queued", "needs_you", "done", "error"}


class FolderIn(BaseModel):
    path: str


class WorkstreamIn(BaseModel):
    folder_id: int
    title: str | None = None
    request: str


class WorkstreamPatch(BaseModel):
    session_id: str | None = None
    status: str | None = None
    note: str | None = None
    title: str | None = None


class RedlineIn(BaseModel):
    path: str
    action: str
    workstream_id: int | None = None
    author: str | None = None


def _desk() -> Desk:
    return service.desk()


def _fail(exc: Exception) -> HTTPException:
    payload = refusal_payload(exc)
    status = 404 if payload["reason"] == "not_found" else 409 if payload.get("refused") else 400
    return HTTPException(status_code=status, detail=f"{payload['message']} {payload.get('hint') or ''}".strip())


def _on_desk(d: Desk, path: str) -> Path:
    resolved = resolve_path(path)
    if d.folder_for_path(resolved) is None:
        raise HTTPException(status_code=403, detail=f"{resolved} is not inside a Desk folder")
    if not resolved.is_file():
        raise HTTPException(status_code=404, detail=f"{resolved} does not exist")
    if resolved.suffix.lower() not in OFFICE_SUFFIXES:
        raise HTTPException(status_code=400, detail="only .docx, .pptx and .xlsx files can be previewed")
    return resolved


def _ws_or_404(d: Desk, ws_id: int) -> dict[str, Any]:
    ws = d.get_workstream(ws_id)
    if not ws:
        raise HTTPException(status_code=404, detail=f"workstream {ws_id} not found")
    return ws


def _rel(d: Desk, path: str) -> str:
    folder = d.folder_for_path(path)
    return str(Path(path).relative_to(folder["path"])) if folder else Path(path).name


@router.get("/health")
def health() -> dict[str, Any]:
    from hermes_office.plugin import engines_ready
    import docx, openpyxl, pptx  # noqa: E401

    return {
        "ok": engines_ready(),
        "version": __version__,
        "engines": {name: getattr(module, "__paper_version__", None) for name, module in (("docx", docx), ("pptx", pptx), ("xlsx", openpyxl))},
        "libreoffice": soffice_path() is not None,
        "data_dir": str(bridge.data_dir()),
    }


@router.get("/desk")
def get_desk() -> dict[str, Any]:
    d = _desk()
    recent = d.list_workstreams(limit=50)
    return {"folders": d.list_folders(), "needs_you": [ws for ws in recent if ws.get("needs_you")], "recent": recent}


@router.post("/folders")
def add_folder(body: FolderIn) -> dict[str, Any]:
    path = resolve_path(body.path)
    if not path.is_dir():
        raise HTTPException(status_code=400, detail=f"{path} is not a folder")
    return _desk().add_folder(path)


@router.delete("/folders/{folder_id}")
def delete_folder(folder_id: int) -> dict[str, Any]:
    d = _desk()
    if d.get_folder(folder_id) is None:
        raise HTTPException(status_code=404, detail=f"folder {folder_id} not found")
    d.remove_folder(folder_id)
    return {"ok": True}


@router.get("/folders/{folder_id}/files")
def folder_files(folder_id: int) -> list[dict[str, Any]]:
    d = _desk()
    folder = d.get_folder(folder_id)
    if folder is None:
        raise HTTPException(status_code=404, detail=f"folder {folder_id} not found")
    root = Path(folder["path"])
    if not root.is_dir():
        raise HTTPException(status_code=410, detail=f"{root} no longer exists")
    touched: dict[str, int] = {}
    for ws in d.list_workstreams(folder_id=folder_id, limit=500):
        for row in d.files(ws["id"]):
            touched[row["path"]] = touched.get(row["path"], 0) + 1
    rows = []
    for file in office_files(root, limit=500):
        stat = file.stat()
        rows.append({"path": str(file), "rel": str(file.relative_to(root)), "kind": kind_of(file), "size": stat.st_size, "mtime": stat.st_mtime, "workstreams": touched.get(str(file.resolve()), 0)})
    return rows


@router.get("/folders/{folder_id}/workstreams")
def folder_workstreams(folder_id: int) -> list[dict[str, Any]]:
    return _desk().list_workstreams(folder_id=folder_id)


@router.post("/workstreams")
def create_workstream(body: WorkstreamIn) -> dict[str, Any]:
    d = _desk()
    if d.get_folder(body.folder_id) is None:
        raise HTTPException(status_code=404, detail=f"folder {body.folder_id} not found")
    title = body.title or service._title_from(body.request, "New request")
    ws = d.create_workstream(body.folder_id, title, request=body.request, status="queued")
    d.add_event(ws["id"], "status", summary="Request queued", data={"request": body.request})
    service.notify(ws["id"], "queued")
    return d.get_workstream(ws["id"])


@router.patch("/workstreams/{ws_id}")
def patch_workstream(ws_id: int, body: WorkstreamPatch) -> dict[str, Any]:
    d = _desk()
    _ws_or_404(d, ws_id)
    fields = body.model_dump(exclude_none=True)
    if "status" in fields:
        if fields["status"] not in STATUSES:
            raise HTTPException(status_code=400, detail=f"status must be one of {sorted(STATUSES)}")
        fields["needs_you"] = 1 if fields["status"] == "needs_you" else 0
    ws = d.update_workstream(ws_id, **fields)
    if "status" in fields:
        d.add_event(ws_id, "status", summary=f"Status → {fields['status']}", data={})
    service.notify(ws_id, ws.get("status"))
    return ws


@router.get("/workstreams/{ws_id}")
def get_workstream(ws_id: int) -> dict[str, Any]:
    d = _desk()
    ws = _ws_or_404(d, ws_id)
    files = [{"path": f["path"], "rel": _rel(d, f["path"]), "kind": f["kind"], "state": f["state"], "receipt_id": f["last_receipt_id"], "exists": Path(f["path"]).exists()} for f in d.files(ws_id)]
    receipts = d.list_receipts(ws_id)
    return {"workstream": ws, "now": d.now(ws_id) if ws.get("status") == "running" else None, "files": files, "events": d.events(ws_id)[-200:], "receipts": receipts}


@router.get("/workstreams/{ws_id}/conversation")
def workstream_conversation(ws_id: int) -> dict[str, Any]:
    ws = _ws_or_404(_desk(), ws_id)
    if not ws.get("session_id"):
        return {"available": False, "reason": "No Hermes session is attached to this workstream yet.", "messages": []}
    return bridge.conversation(ws["session_id"])


def _baseline(d: Desk, path: Path, receipt_id: int | None) -> dict[str, Any]:
    receipt = d.get_receipt(receipt_id) if receipt_id is not None else d.first_receipt_for(path)
    if receipt is None or receipt["path"] != str(path):
        raise HTTPException(status_code=404, detail=f"no receipt with a backup for {path.name}")
    if not receipt.get("backup_path") or not Path(receipt["backup_path"]).exists():
        raise HTTPException(status_code=404, detail=f"{path.name} was created by this workstream; there is no before version")
    return receipt


@router.get("/preview")
def preview(path: str, receipt_id: int | None = None) -> dict[str, Any]:
    d = _desk()
    resolved = _on_desk(d, path)
    try:
        if receipt_id is None:
            return preview_file(resolved)
        receipt = _baseline(d, resolved, receipt_id)
        return {**preview_file(receipt["backup_path"]), "path": str(resolved), "baseline_of": receipt["id"]}
    except HTTPException:
        raise
    except Exception as exc:
        raise _fail(exc) from exc


@router.get("/diff")
def diff(path: str, receipt_id: int | None = None) -> dict[str, Any]:
    d = _desk()
    resolved = _on_desk(d, path)
    receipt = _baseline(d, resolved, receipt_id)
    try:
        return {**diff_file(resolved, receipt["backup_path"]), "baseline_receipt": receipt["id"]}
    except Exception as exc:
        raise _fail(exc) from exc


@router.post("/redline")
def redline(body: RedlineIn) -> dict[str, Any]:
    d = _desk()
    resolved = _on_desk(d, body.path)
    if body.action not in ("accept", "reject"):
        raise HTTPException(status_code=400, detail="action must be accept or reject")
    try:
        out = service.resolve_redline(str(resolved), action=body.action, workstream_id=body.workstream_id, author=body.author, scope="author" if body.author else "all", d=d)
    except Exception as exc:
        raise _fail(exc) from exc
    receipt = d.get_receipt(out["receipt_id"]) if out["receipt_id"] else None
    return {"ok": True, "resolved": out["result"]["resolved"], "remaining": out["result"]["remaining"], "receipt": receipt}


@router.post("/receipts/{receipt_id}/revert")
def revert(receipt_id: int) -> dict[str, Any]:
    d = _desk()
    receipt = d.get_receipt(receipt_id)
    if receipt is None:
        raise HTTPException(status_code=404, detail=f"receipt {receipt_id} not found")
    if d.folder_for_path(receipt["path"]) is None:
        raise HTTPException(status_code=403, detail=f"{receipt['path']} is not inside a Desk folder")
    try:
        new = service.revert_receipt(receipt_id, d=d)
    except (OfficeError, OSError) as exc:
        raise _fail(exc) from exc
    return {"ok": True, "receipt": new}


@router.get("/events")
def events(since: str | None = None) -> list[dict[str, Any]]:
    return _desk().events_since(since)
