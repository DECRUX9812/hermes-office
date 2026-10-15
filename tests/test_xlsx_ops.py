from __future__ import annotations

import json
from pathlib import Path

from hermes_office.desk import Desk
from hermes_office.guard import sha256
from hermes_office.tools import xlsx_edit_handler, xlsx_recalc_handler
from hermes_office.xlsx_ops import diff_xlsx, inspect_xlsx, preview_xlsx, recalc_xlsx
from openpyxl import load_workbook


def test_xlsx_edit_recalc_and_diff(sample_files: dict[str, Path], desk: Desk) -> None:
    path = sample_files["xlsx"]
    before = sha256(path)
    result = json.loads(xlsx_edit_handler({"path": str(path), "ops": [{"kind": "insert_rows", "sheet": "Budget", "index": 1, "amount": 1}]}))
    assert result["ok"] is True
    assert sha256(path) != before
    wb = load_workbook(path)
    assert wb["Budget"]["C2"].value == "=A2+B2"

    refused_before = sha256(path)
    refused = json.loads(xlsx_edit_handler({"path": str(path), "ops": [{"kind": "rename_sheet", "sheet": "Missing", "new_name": "New"}]}))
    assert refused["refused"] is True
    assert sha256(path) == refused_before
    assert any(event["kind"] == "refused" for event in desk.events(desk.list_workstreams()[0]["id"]))

    degraded = recalc_xlsx(path, read=["Budget!C2"], soffice=None)
    assert degraded["recalculated"] is False
    assert degraded["reason"] == "libreoffice_missing"

    routed = json.loads(xlsx_recalc_handler({"path": str(path), "read": ["Budget!C2"]}))
    assert routed["ok"] is True

    assert diff_xlsx(sample_files["xlsx"], sample_files["xlsx"])["changed"] == []
    assert inspect_xlsx(sample_files["xlsx"])["kind"] == "xlsx"
    assert preview_xlsx(sample_files["xlsx"])["xlsx"]["sheets"]
