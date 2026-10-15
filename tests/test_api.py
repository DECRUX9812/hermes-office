from __future__ import annotations

import importlib.util
from pathlib import Path

import pytest
from fastapi import FastAPI
from fastapi.testclient import TestClient

from conftest import REPO, call
from hermes_office.guard import sha256


@pytest.fixture
def client(hermes_home) -> TestClient:
    # Load the router the way Hermes does: by file path.
    spec = importlib.util.spec_from_file_location("hermes_dashboard_plugin_hermes_office", REPO / "dashboard" / "plugin_api.py")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    app = FastAPI()
    app.include_router(module.router, prefix="/api/plugins/hermes-office")
    return TestClient(app)


API = "/api/plugins/hermes-office"


def test_health(client) -> None:
    body = client.get(f"{API}/health").json()
    assert body["ok"] is True
    assert body["engines"]["docx"] and body["engines"]["pptx"] and body["engines"]["xlsx"]
    assert body["data_dir"].endswith("plugin-data/hermes-office")


def test_folders_files_and_desk(client, files, tmp_path) -> None:
    desk = client.get(f"{API}/desk").json()
    folder = desk["folders"][0]
    assert folder["name"] == "Client Work" and folder["counts"] == {"files": 0, "workstreams": 0, "needs_you": 0}
    listed = client.get(f"{API}/folders/{folder['id']}/files").json()
    assert {row["rel"] for row in listed} >= {"Services Agreement.docx", "Campaign Budget FY26.xlsx", "Spring Campaign Review.pptx"}
    assert client.post(f"{API}/folders", json={"path": str(tmp_path / "missing")}).status_code == 400
    other = tmp_path / "Other"
    other.mkdir()
    added = client.post(f"{API}/folders", json={"path": str(other)}).json()
    assert added["path"] == str(other.resolve())
    assert client.delete(f"{API}/folders/{added['id']}").json() == {"ok": True}
    assert client.delete(f"{API}/folders/{added['id']}").status_code == 404


def test_new_request_lifecycle(client, files) -> None:
    folder_id = client.get(f"{API}/desk").json()["folders"][0]["id"]
    ws = client.post(f"{API}/workstreams", json={"folder_id": folder_id, "request": "Move payment terms to 45 days"}).json()
    assert ws["status"] == "queued" and ws["title"] == "Move payment terms to 45 days"
    patched = client.patch(f"{API}/workstreams/{ws['id']}", json={"session_id": "chat-x", "status": "running"}).json()
    assert patched["session_id"] == "chat-x" and patched["status"] == "running"
    assert client.patch(f"{API}/workstreams/{ws['id']}", json={"status": "bogus"}).status_code == 400
    call("docx_tracked_edit", {"path": str(files["docx"]), "find": "within 30 days", "replace": "within 45 days"}, session="chat-x")
    desk = client.get(f"{API}/desk").json()
    assert [w["id"] for w in desk["needs_you"]] == [ws["id"]]
    detail = client.get(f"{API}/workstreams/{ws['id']}").json()
    assert detail["files"][0]["rel"] == "Services Agreement.docx" and detail["files"][0]["state"] == "changed"
    assert [e["kind"] for e in detail["events"]][-2:] == ["tool_started", "tool_finished"]
    assert detail["receipts"][0]["tool"] == "docx_tracked_edit"
    convo = client.get(f"{API}/workstreams/{ws['id']}/conversation").json()
    assert convo["available"] is False and convo["reason"]


def test_preview_before_after_and_diff(client, files) -> None:
    path = files["docx"]
    out = call("docx_tracked_edit", {"path": str(path), "find": "within 30 days", "replace": "within 45 days"})
    after = client.get(f"{API}/preview", params={"path": str(path)}).json()
    runs = [run for p in after["docx"]["paragraphs"] for run in p["runs"]]
    assert {"text": "30", "rev": "del", "author": "Hermes"} in runs
    assert {"text": "45", "rev": "ins", "author": "Hermes"} in runs
    payment = next(p for p in after["docx"]["paragraphs"] if any(r["text"] == "45" for r in p["runs"]))
    assert "".join(r["text"] for r in payment["runs"]).startswith("The Client will pay the fees in Schedule B. Payment is due within 3045 days")
    before = client.get(f"{API}/preview", params={"path": str(path), "receipt_id": out["receipt_id"]}).json()
    assert all(r["rev"] is None for p in before["docx"]["paragraphs"] for r in p["runs"])
    diff = client.get(f"{API}/diff", params={"path": str(path)}).json()
    assert diff["kind"] == "docx"
    assert diff["paragraphs_changed"] == [{"before": "The Client will pay the fees in Schedule B. Payment is due within 30 days of the invoice date.", "after": "The Client will pay the fees in Schedule B. Payment is due within 45 days of the invoice date."}]
    assert {r["revision_type"] for r in diff["revisions"]} == {"insertion", "deletion"}


def test_preview_xlsx_and_pptx_shapes(client, files) -> None:
    xlsx = client.get(f"{API}/preview", params={"path": str(files["xlsx"])}).json()
    budget = xlsx["xlsx"]["sheets"][0]
    assert budget["rows"][0][0] == {"ref": "A1", "value": "Line item", "formula": None}
    assert budget["rows"][1][5]["formula"] == "=SUM(B2:E2)"
    pptx = client.get(f"{API}/preview", params={"path": str(files["pptx"])}).json()["pptx"]
    assert pptx["slide_w"] == 9144000 and len(pptx["slides"]) == 3
    assert pptx["slides"][0]["title"] == "Spring Campaign Review"
    revenue = next(s for s in pptx["slides"][1]["shapes"] if s["name"] == "Revenue")
    assert revenue["kind"] == "TEXT_BOX" and revenue["text"].startswith("Revenue grew 12%")


def test_xlsx_and_pptx_diff_against_receipt(client, files) -> None:
    call("xlsx_edit", {"path": str(files["xlsx"]), "ops": [{"op": "set", "ref": "Budget!C4", "number": 13500}]})
    xdiff = client.get(f"{API}/diff", params={"path": str(files["xlsx"])}).json()
    assert xdiff["kind"] == "xlsx" and xdiff["changed"][0]["address"] == "Budget!C4"
    call("pptx_edit_shape", {"path": str(files["pptx"]), "slide": 1, "shape": "Revenue", "ops": [{"op": "replace_text", "find": "12%", "replace": "15%"}]})
    pdiff = client.get(f"{API}/diff", params={"path": str(files["pptx"])}).json()
    assert pdiff["kind"] == "pptx" and "15%" in str(pdiff["slide_changes"])


def test_preview_refuses_files_off_the_desk(client, files, outside) -> None:
    assert client.get(f"{API}/preview", params={"path": str(outside["docx"])}).status_code == 403
    assert client.get(f"{API}/preview", params={"path": str(files["folder"] / "nope.docx")}).status_code == 404
    assert client.get(f"{API}/diff", params={"path": str(files["docx"])}).status_code == 404


def test_redline_accept_from_needs_you(client, files) -> None:
    out = call("docx_tracked_edit", {"path": str(files["docx"]), "find": "within 30 days", "replace": "within 45 days"})
    accepted = client.post(f"{API}/redline", json={"path": str(files["docx"]), "action": "accept", "workstream_id": out["workstream_id"]}).json()
    assert accepted["ok"] is True and accepted["resolved"] == 2 and accepted["remaining"] == 0
    assert accepted["receipt"]["tool"] == "docx_redline_accept_reject"
    detail = client.get(f"{API}/workstreams/{out['workstream_id']}").json()
    assert detail["workstream"]["status"] == "done" and detail["workstream"]["needs_you"] == 0
    assert client.post(f"{API}/redline", json={"path": str(files["docx"]), "action": "maybe"}).status_code == 400


def test_revert_route_and_events(client, files) -> None:
    path = files["xlsx"]
    original = sha256(path)
    out = call("xlsx_edit", {"path": str(path), "ops": [{"op": "set", "ref": "Budget!C4", "number": 13500}]})
    reverted = client.post(f"{API}/receipts/{out['receipt_id']}/revert").json()
    assert reverted["ok"] is True and sha256(path) == original
    assert client.post(f"{API}/receipts/{out['receipt_id']}/revert").status_code == 404
    events = client.get(f"{API}/events").json()
    assert events[-1]["summary"] == f"Reverted receipt {out['receipt_id']}"
    assert client.get(f"{API}/events", params={"since": events[-1]["ts"]}).json() == []
