from __future__ import annotations

import json
import os
import shutil
import subprocess
import sys
from pathlib import Path

import pytest

from conftest import REPO

VENV_PYTHON = Path("/home/user/Code/ship-test/office-venv/.venv/bin/python")


def _venv() -> Path:
    if not VENV_PYTHON.exists():
        pytest.skip("office venv is not built")
    return VENV_PYTHON


@pytest.fixture
def venv_home(tmp_path: Path) -> Path:
    """An isolated HERMES_HOME with hermes-office installed as a directory plugin and enabled."""
    home = tmp_path / "hermes-home"
    plugins = home / "plugins" / "hermes-office"
    plugins.parent.mkdir(parents=True)
    shutil.copytree(REPO, plugins, ignore=shutil.ignore_patterns(".git", ".task", "tests", "__pycache__", "desktop"))
    (home / "config.yaml").write_text("plugins:\n  enabled:\n    - hermes-office\n", encoding="utf-8")
    return home


def _run(home: Path, code: str) -> subprocess.CompletedProcess:
    env = os.environ.copy()
    env["HERMES_HOME"] = str(home)
    env.pop("HERMES_PROFILE", None)
    return subprocess.run([str(_venv()), "-c", code], capture_output=True, text=True, env=env, timeout=300)


def test_real_hermes_discovers_all_14_tools_and_the_skill(venv_home: Path) -> None:
    script = """
import json
from hermes_cli.plugins import discover_plugins, get_plugin_manager
discover_plugins(force=True)
manager = get_plugin_manager()
plugin = manager._plugins.get("hermes-office")
from tools.registry import registry
names = ["office_desk", "office_inspect", "office_search", "docx_tracked_edit", "docx_comment",
         "docx_redline_accept_reject", "pptx_inspect_formats", "pptx_edit_shape", "pptx_compose_decks",
         "pptx_diff", "xlsx_edit", "xlsx_recalc", "xlsx_diff", "office_receipt"]
missing = [name for name in names if registry.get_entry(name) is None]
print("RESULT:" + json.dumps({
    "plugin_error": getattr(plugin, "error", None),
    "missing_tools": missing,
    "skill": manager.list_plugin_skills("hermes-office"),
}))
"""
    result = _run(venv_home, script)
    assert result.returncode == 0, result.stderr[-3000:]
    lines = [line for line in (result.stdout + result.stderr).splitlines() if line.startswith("RESULT:")]
    assert lines, f"no RESULT in stdout:\n{result.stdout[-2000:]}\nstderr:\n{result.stderr[-2000:]}"
    report = json.loads(lines[-1].removeprefix("RESULT:"))
    assert report["plugin_error"] is None, report
    assert report["missing_tools"] == [], report
    assert report["skill"] == ["office-review"], report


def test_tools_dispatch_through_the_real_registry(venv_home: Path, tmp_path: Path) -> None:
    files = tmp_path / "client-work"
    files.mkdir()
    demo = subprocess.run([str(_venv()), str(REPO / "scripts" / "make_demo_files.py"), str(files)], capture_output=True, text=True, check=True)
    assert demo.returncode == 0
    target = files / "Services Agreement.docx"
    script = f"""
import json
from hermes_cli.plugins import discover_plugins
discover_plugins(force=True)
from tools.registry import registry
raw = registry.dispatch("office_desk", {{"action": "add_folder", "path": {str(files)!r}}}, session_id="chat-1")
add = json.loads(raw)
raw = registry.dispatch("docx_tracked_edit", {{"path": {str(target)!r}, "find": "within 30 days", "replace": "within 45 days"}}, session_id="chat-1")
edit = json.loads(raw)
import docx
revisions = [(r.revision_type, r.text) for r in docx.Document({str(target)!r}).revisions]
raw = registry.dispatch("office_receipt", {{"action": "list"}}, session_id="chat-1")
receipts = json.loads(raw)
print("RESULT:" + json.dumps({{"add_ok": add["ok"], "edit_ok": edit["ok"], "needs_you": edit.get("receipt_id") is not None,
    "revisions": revisions, "receipts": len(receipts["receipts"])}}))
"""
    result = _run(venv_home, script)
    assert result.returncode == 0, result.stderr[-3000:]
    lines = [line for line in (result.stdout + result.stderr).splitlines() if line.startswith("RESULT:")]
    assert lines, f"no RESULT in stdout:\n{result.stdout[-2000:]}\nstderr:\n{result.stderr[-2000:]}"
    report = json.loads(lines[-1].removeprefix("RESULT:"))
    assert report["add_ok"] and report["edit_ok"], report
    assert ("insertion", "45") in [tuple(r) for r in report["revisions"]], report
    assert report["receipts"] >= 1
    # Refusal path through the real dispatcher leaves the file byte-identical.
    target = files / "Services Agreement (Client edits).docx"  # untouched
    before = target.read_bytes()
    script = f"""
import json
from hermes_cli.plugins import discover_plugins
discover_plugins(force=True)
from tools.registry import registry
raw = registry.dispatch("docx_tracked_edit", {{"path": {str(target)!r}, "find": "days", "replace": "x"}}, session_id="chat-1")
print("RESULT:" + raw)
"""
    result = _run(venv_home, script)
    assert result.returncode == 0, result.stderr[-3000:]
    lines = [line for line in (result.stdout + result.stderr).splitlines() if line.startswith("RESULT:")]
    assert lines, f"no RESULT in stdout:\n{result.stdout[-2000:]}"
    refused = json.loads(lines[-1].removeprefix("RESULT:"))
    assert refused["ok"] is False and refused["refused"] is True and refused["reason"] == "AmbiguousTargetError", refused
    assert target.read_bytes() == before


def test_demo_file_generator(venv_home: Path, tmp_path: Path) -> None:
    out = tmp_path / "demo"
    result = subprocess.run([str(_venv()), str(REPO / "scripts" / "make_demo_files.py"), str(out)], capture_output=True, text=True)
    assert result.returncode == 0, result.stderr
    for name in ("Services Agreement.docx", "Services Agreement (Client edits).docx", "Campaign Budget FY26.xlsx",
                 "Spring Campaign Review.pptx", "Spring Campaign Review v2.pptx", "Appendix.pptx"):
        assert (out / name).is_file(), name
