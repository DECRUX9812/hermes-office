from __future__ import annotations

from pathlib import Path

from hermes_office.desk import Desk
from hermes_office.guard import atomic_write, backup, kind_of, resolve_path, sha256


def test_guard_helpers(tmp_path: Path) -> None:
    target = tmp_path / "x.txt"
    target.write_text("hello", encoding="utf-8")
    assert resolve_path(target) == target.resolve()
    assert kind_of(tmp_path / "a.docx") == "docx"
    assert kind_of(tmp_path / "a.pptx") == "pptx"
    assert kind_of(tmp_path / "a.xlsx") == "xlsx"
    assert sha256(target)

    out = atomic_write(target, lambda tmp: tmp.write_text("bye", encoding="utf-8"))
    assert out.read_text(encoding="utf-8") == "bye"

    d = Desk(tmp_path / "data")
    folder = tmp_path / "folder"
    folder.mkdir()
    added = d.add_folder(folder)
    ws = d.create_workstream(added["id"], "test")
    assert d.folder_for_path(folder / "file.docx")["id"] == added["id"]
    d.track_file(ws["id"], folder / "file.docx", "changed", kind="docx")
    receipt = d.add_receipt(ws["id"], "tool", folder / "file.docx", "op", None, None, None, {"ok": True})
    assert d.get_receipt(receipt["id"])["tool"] == "tool"
    assert d.list_folders()[0]["counts"]["workstreams"] == 1
    assert d.files(ws["id"])[0]["state"] == "changed"


def test_backup_writes_copy(tmp_path: Path) -> None:
    data_dir = tmp_path / "data"
    source = tmp_path / "source.docx"
    source.write_text("content", encoding="utf-8")
    copied = backup(data_dir, 3, 9, source)
    assert copied.read_text(encoding="utf-8") == "content"
