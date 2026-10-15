from __future__ import annotations

import json
from pathlib import Path

from fastapi import FastAPI
from fastapi.testclient import TestClient

from hermes_office.plugin import register
from hermes_office.tools import docx_tracked_edit, office_desk, office_receipt_handler


class Recorder:
    def __init__(self) -> None:
        self.tools = {}
        self.skills = {}

    def register_tool(self, **kwargs):
        self.tools[kwargs["name"]] = kwargs

    def register_skill(self, name, path):
        self.skills[name] = Path(path)


def test_registers_tools_and_skill() -> None:
    ctx = Recorder()
    register(ctx)
    assert len(ctx.tools) == 14
    assert "office-review" in ctx.skills


def test_plugin_api_routes(sample_files: dict[str, Path], hermes_home: Path) -> None:
    from dashboard.plugin_api import router

    app = FastAPI()
    app.include_router(router)
    client = TestClient(app)

    health = client.get("/health")
    assert health.status_code == 200
    assert health.json()["ok"] is True

    folder = client.post("/folders", json={"path": str(sample_files["folder"])})
    assert folder.status_code == 200
    folder_id = folder.json()["id"]

    files = client.get(f"/folders/{folder_id}/files")
    assert files.status_code == 200
    assert any(item["kind"] == "docx" for item in files.json())

    preview = client.get("/preview", params={"path": str(sample_files["docx"])})
    assert preview.status_code == 200
    assert preview.json()["kind"] == "docx"

    diff = client.get("/diff", params={"path": str(sample_files["docx"])})
    assert diff.status_code in {200, 404}

    ws = client.post("/workstreams", json={"folder_id": folder_id, "title": "Fix doc", "request": "edit"})
    assert ws.status_code == 200
    ws_id = ws.json()["id"]
    status = client.get(f"/workstreams/{ws_id}")
    assert status.status_code == 200


def test_desk_and_receipt_tools(sample_files: dict[str, Path], desk) -> None:
    folder = sample_files["folder"]
    added = json.loads(office_desk({"action": "add_folder", "path": str(folder)}))
    assert added["ok"] is True
    edit = json.loads(docx_tracked_edit({"path": str(sample_files["docx"]), "find": "Payment is due within 30 days of the invoice date.", "replace": "Payment is due within 14 days of the invoice date.", "nth": 1, "author": "A"}))
    receipt_id = edit["receipt"]["id"]
    shown = json.loads(office_receipt_handler({"action": "show", "receipt_id": receipt_id}))
    assert shown["ok"] is True
    reverted = json.loads(office_receipt_handler({"action": "revert", "receipt_id": receipt_id}))
    assert reverted["ok"] is True
    assert json.loads(office_desk({"action": "needs_you", "workstream_id": edit["receipt"]["workstream_id"], "note": "review"}))["ok"] is True
    assert json.loads(office_desk({"action": "done", "workstream_id": edit["receipt"]["workstream_id"]}))["ok"] is True

