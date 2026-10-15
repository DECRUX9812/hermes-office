from __future__ import annotations

import json
import os
import shutil
import subprocess
import sys
from pathlib import Path

import pytest


def test_make_demo_files(tmp_path: Path) -> None:
    out = tmp_path / "demo"
    subprocess.run([sys.executable, "scripts/make_demo_files.py", str(out)], check=True)
    assert (out / "Services Agreement.docx").exists()
    assert (out / "Budget FY26.xlsx").exists()
    assert (out / "Quarterly Review.pptx").exists()
    assert (out / "Quarterly Review v2.pptx").exists()


def test_real_hermes_discovery(tmp_path: Path, hermes_home: Path) -> None:
    repo = Path.cwd()
    plugin_home = hermes_home / "plugins" / "hermes-office"
    plugin_home.parent.mkdir(parents=True, exist_ok=True)
    if hasattr(os, "symlink"):
        os.symlink(repo, plugin_home)
    else:
        shutil.copytree(repo, plugin_home)
    (hermes_home / "config.yaml").write_text("plugins:\n  enabled:\n    - hermes-office\n", encoding="utf-8")

    py = Path("/home/user/Code/ship-test/office-venv/.venv/bin/python")
    probe_env = os.environ.copy()
    probe_env["HERMES_HOME"] = str(hermes_home)
    probe = subprocess.run([str(py), "-c", "import hermes_cli.plugins"], capture_output=True, text=True, env=probe_env)
    if probe.returncode != 0 and "ruamel" in probe.stderr:
        pytest.skip("upstream hermes discovery import is unavailable in this venv (missing ruamel)")
    script = """
from hermes_cli.plugins import discover_plugins
from tools.registry import registry
discover_plugins(force=True)
names = ["office_desk", "office_inspect", "office_search", "docx_tracked_edit", "docx_comment", "docx_redline_accept_reject", "pptx_inspect_formats", "pptx_edit_shape", "pptx_compose_decks", "pptx_diff", "xlsx_edit", "xlsx_recalc", "xlsx_diff", "office_receipt"]
print(json.dumps({name: registry.get_entry(name) is not None for name in names}))
"""
    env = os.environ.copy()
    env["HERMES_HOME"] = str(hermes_home)
    result = subprocess.run([str(py), "-c", "import json;" + script], check=True, capture_output=True, text=True, env=env)
    data = json.loads(result.stdout.strip())
    assert all(data.values())
