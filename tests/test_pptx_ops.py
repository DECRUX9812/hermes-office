from __future__ import annotations

import json
from pathlib import Path

from hermes_office.desk import Desk
from hermes_office.guard import sha256
from hermes_office.pptx_ops import compose_pptx_decks, diff_pptx, edit_pptx_shape, inspect_pptx, inspect_pptx_formats, preview_pptx
from hermes_office.tools import pptx_compose_decks_handler, pptx_edit_shape_handler, pptx_inspect_formats_handler


def test_pptx_edit_compose_and_diff(sample_files: dict[str, Path], desk: Desk) -> None:
    path = sample_files["pptx"]
    before = sha256(path)
    result = json.loads(pptx_edit_shape_handler({"path": str(path), "slide": 1, "shape": "metric-box", "ops": [{"kind": "set_geometry", "x": 100, "y": 100, "w": 200, "h": 200}]}))
    assert result["ok"] is True
    assert sha256(path) != before

    refused_before = sha256(path)
    refused = json.loads(pptx_edit_shape_handler({"path": str(path), "slide": 1, "shape": "missing", "ops": [{"kind": "delete"}]}))
    assert refused["refused"] is True
    assert sha256(path) == refused_before
    assert any(event["kind"] == "refused" for event in desk.events(desk.list_workstreams()[0]["id"]))

    composed = tmp = sample_files["folder"] / "composed.pptx"
    ok = json.loads(pptx_compose_decks_handler({"dest": str(path), "sources": [{"path": str(sample_files["pptx"]), "slides": [0, 1]}], "mode": "keep_appearance", "output_path": str(composed)}))
    assert ok["ok"] is True
    assert composed.exists()

    diff = diff_pptx(sample_files["pptx"], sample_files["pptx_v2"])
    assert diff["schema"] == "paper-deck-diff"
    assert preview_pptx(sample_files["pptx"])["pptx"]["slides"]
    assert inspect_pptx(sample_files["pptx"])["kind"] == "pptx"
    assert inspect_pptx_formats(sample_files["pptx"], slide=0, shape="hero-box")["shape"]
