#!/usr/bin/env python3
"""Set up a rich demo Desk state with generated fixture files, workstreams, tracked edits, and receipts."""

from __future__ import annotations

import json
import os
import shutil
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO))


def setup_desk(home_dir: Path, desk_dir: Path) -> dict:
    home_dir.mkdir(parents=True, exist_ok=True)
    desk_dir.mkdir(parents=True, exist_ok=True)

    # 1. Generate clean demo files
    from scripts.make_demo_files import make_all
    files = make_all(desk_dir)

    # 2. Configure environment
    os.environ["HERMES_HOME"] = str(home_dir)
    os.environ.pop("HERMES_PROFILE", None)

    import hermes_office.service as service
    from hermes_office import tools

    desk = service.desk()

    # Register folder
    folder = desk.add_folder(str(desk_dir))

    # Workstream 1: "Review Services Agreement Redlines" (needs_you)
    ws1 = desk.create_workstream(
        folder["id"],
        request="Review the client redlines on payment terms and liability cap in the Services Agreement.",
        title="Services Agreement Legal Review",
        session_id="session-legal-rev-101",
        status="needs_you",
    )
    desk.update_workstream(ws1["id"], note="Client modified payment window from 30 to 45 days. Review redline.", needs_you=True)

    # Apply tracked edit via tool
    r1 = tools.docx_tracked_edit({
        "path": str(files["docx"]),
        "find": "within 30 days of the invoice date",
        "replace": "within 45 days of the invoice date",
        "author": "Client Legal Counsel",
        "workstream_id": ws1["id"],
    }, session_id="session-legal-rev-101")

    # Workstream 2: "FY26 Campaign Budget Recalculation" (running)
    ws2 = desk.create_workstream(
        folder["id"],
        request="Insert new Events line item, update Media spend, and recalculate all Q1-Q4 formulas with LibreOffice.",
        title="FY26 Budget Modeling & Recalc",
        session_id="session-budget-202",
        status="running",
    )
    desk.update_workstream(ws2["id"], note="LibreOffice formula oracle running...")

    r2 = tools.xlsx_edit({
        "path": str(files["xlsx"]),
        "ops": [
            {"op": "set", "ref": "Budget!C4", "number": 13500},
            {"op": "insert_rows", "sheet": "Budget", "index": 6, "amount": 1},
            {"op": "set", "ref": "Budget!A6", "value": "Events"},
            {"op": "set", "ref": "Budget!B6", "number": 1000},
            {"op": "formula", "ref": "Budget!F6", "value": "SUM(B6:E6)"},
        ],
        "workstream_id": ws2["id"],
    }, session_id="session-budget-202")

    r2_calc = tools.xlsx_recalc({
        "path": str(files["xlsx"]),
        "read": ["Budget!C4", "Budget!F6", "Budget!F7", "Assumptions!B3"],
        "workstream_id": ws2["id"],
    }, session_id="session-budget-202")

    # Workstream 3: "Spring Campaign Pitch Deck Review" (done)
    ws3 = desk.create_workstream(
        folder["id"],
        request="Diff the review deck against v2, inspect slide typography provenance, and compose the Appendix.",
        title="Pitch Deck Deck Diff & Composition",
        session_id="session-pitch-303",
        status="done",
    )
    desk.update_workstream(ws3["id"], note="Deck diff generated and Appendix composed.")

    r3_edit = tools.pptx_edit_shape({
        "path": str(files["pptx"]),
        "slide": 1,
        "shape": "Revenue",
        "ops": [{"op": "replace_text", "find": "12%", "replace": "15%"}],
        "workstream_id": ws3["id"],
    }, session_id="session-pitch-303")

    r3_comp = tools.pptx_compose_decks({
        "path": str(files["pptx"]),
        "sources": [{"path": str(files["appendix"]), "slides": [0]}],
        "mode": "adopt_theme",
        "output_path": str(desk_dir / "Spring Campaign Combined.pptx"),
        "workstream_id": ws3["id"],
    }, session_id="session-pitch-303")

    return {
        "folder": folder,
        "workstreams": [ws1, ws2, ws3],
        "receipts": desk.list_receipts(),
    }


if __name__ == "__main__":
    home = Path("/tmp/office-demo-home")
    desk_f = Path("/tmp/office-demo-folder")
    shutil.rmtree(home, ignore_errors=True)
    shutil.rmtree(desk_f, ignore_errors=True)
    res = setup_desk(home, desk_f)
    print(f"Setup complete: folder={res['folder']['name']}, workstreams={len(res['workstreams'])}, receipts={len(res['receipts'])}")
