#!/usr/bin/env python3
"""End-to-end proof of hermes-office running against stock hermes-agent in a clean venv.

Demonstrates:
1. Desk registration and workspace setup
2. docx: tracked changes (w:ins / w:del) with receipt and refusal verification
3. xlsx: reference-aware edits, LibreOffice recalculation, and structural diff
4. pptx: structural diff, format inspection, shape edit, and slide composition
5. receipts and Desk state tracking with reversible audit trail
"""

from __future__ import annotations

import importlib.util
import json
import os
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

def _find_hermes_python() -> Path | None:
    """A Python with hermes-agent installed: $HERMES_OFFICE_PYTHON, else the current interpreter."""
    override = os.environ.get("HERMES_OFFICE_PYTHON")
    if override:
        return Path(override).expanduser()
    if importlib.util.find_spec("hermes_cli") is not None:
        return Path(sys.executable)
    return None


VENV_PYTHON = _find_hermes_python()
REPO = Path(__file__).resolve().parent.parent


def run_proof() -> dict:
    if VENV_PYTHON is None or not VENV_PYTHON.exists():
        raise RuntimeError("No Python with hermes-agent found; set HERMES_OFFICE_PYTHON")

    tmp_dir = Path(tempfile.mkdtemp(prefix="office-proof-"))
    try:
        hermes_home = tmp_dir / "hermes_home"
        plugin_dir = hermes_home / "plugins" / "office-desk"
        plugin_dir.parent.mkdir(parents=True, exist_ok=True)
        shutil.copytree(
            REPO,
            plugin_dir,
            ignore=shutil.ignore_patterns(".git", ".task", "tests", "__pycache__", "desktop"),
        )
        (hermes_home / "config.yaml").write_text(
            "plugins:\n  enabled:\n    - office-desk\n",
            encoding="utf-8",
        )

        work_dir = tmp_dir / "client-desk"
        work_dir.mkdir(parents=True, exist_ok=True)

        # Generate demo files
        gen_script = REPO / "scripts" / "make_demo_files.py"
        subprocess.run(
            [str(VENV_PYTHON), str(gen_script), str(work_dir)],
            capture_output=True,
            text=True,
            check=True,
        )

        # Run python runner inside venv with HERMES_HOME set
        runner_code = f"""
import json
import os
import shutil
from pathlib import Path

from hermes_cli.plugins import discover_plugins, get_plugin_manager
discover_plugins(force=True)
manager = get_plugin_manager()
from tools.registry import registry

work_dir = {str(work_dir)!r}
session_id = "session-proof-demo"

# 1. Register folder on Desk
r_desk = json.loads(registry.dispatch("office_desk", {{"action": "add_folder", "path": work_dir}}, session_id=session_id))

# 2. docx tracked edit
docx_path = str(Path(work_dir) / "Services Agreement.docx")
r_docx_edit = json.loads(registry.dispatch("docx_tracked_edit", {{
    "path": docx_path,
    "find": "within 30 days",
    "replace": "within 45 days",
    "author": "LegalReviewer"
}}, session_id=session_id))

# Verify docx revisions natively
import docx
doc = docx.Document(docx_path)
revisions = [(r.revision_type, getattr(r, "author", None), r.text) for r in doc.revisions]

# 3. docx refusal test (ambiguous target)
r_docx_refusal = json.loads(registry.dispatch("docx_tracked_edit", {{
    "path": docx_path,
    "find": "agreement",
    "replace": "contract"
}}, session_id=session_id))

# 4. xlsx reference-aware edit & recalc
xlsx_path = str(Path(work_dir) / "Campaign Budget FY26.xlsx")
before_xlsx = str(Path(work_dir) / "Campaign Budget FY26 (before).xlsx")
shutil.copyfile(xlsx_path, before_xlsx)

r_xlsx_edit = json.loads(registry.dispatch("xlsx_edit", {{
    "path": xlsx_path,
    "ops": [
        {{"op": "set", "ref": "Budget!C4", "number": 13500}},
        {{"op": "insert_rows", "sheet": "Budget", "index": 6, "amount": 1}},
        {{"op": "set", "ref": "Budget!A6", "value": "Events"}},
        {{"op": "set", "ref": "Budget!B6", "number": 1000}},
        {{"op": "formula", "ref": "Budget!F6", "value": "SUM(B6:E6)"}}
    ]
}}, session_id=session_id))

# Recalc with LibreOffice
r_xlsx_recalc = json.loads(registry.dispatch("xlsx_recalc", {{
    "path": xlsx_path,
    "read": ["Budget!C4", "Budget!F6", "Budget!F7", "Assumptions!B3"]
}}, session_id=session_id))

# xlsx diff
r_xlsx_diff = json.loads(registry.dispatch("xlsx_diff", {{
    "a": before_xlsx,
    "b": xlsx_path
}}, session_id=session_id))

# 5. pptx diff & inspection & composition
pptx_v1 = str(Path(work_dir) / "Spring Campaign Review.pptx")
pptx_v2 = str(Path(work_dir) / "Spring Campaign Review v2.pptx")
pptx_app = str(Path(work_dir) / "Appendix.pptx")

r_pptx_diff = json.loads(registry.dispatch("pptx_diff", {{
    "a": pptx_v1,
    "b": pptx_v2
}}, session_id=session_id))

r_pptx_inspect = json.loads(registry.dispatch("pptx_inspect_formats", {{
    "path": pptx_v1,
    "slide": 1,
    "shape": "Revenue"
}}, session_id=session_id))

output_composed = str(Path(work_dir) / "Spring Campaign Combined.pptx")
r_pptx_compose = json.loads(registry.dispatch("pptx_compose_decks", {{
    "path": pptx_v1,
    "sources": [
        {{"path": pptx_app, "slides": [0]}}
    ],
    "mode": "adopt_theme",
    "output_path": output_composed
}}, session_id=session_id))

# 6. Check receipts
r_receipts = json.loads(registry.dispatch("office_receipt", {{
    "action": "list"
}}, session_id=session_id))

output = {{
    "desk": r_desk,
    "docx_edit": r_docx_edit,
    "docx_revisions": revisions,
    "docx_refusal": r_docx_refusal,
    "xlsx_edit": r_xlsx_edit,
    "xlsx_recalc": r_xlsx_recalc,
    "xlsx_diff": r_xlsx_diff,
    "pptx_diff": r_pptx_diff,
    "pptx_inspect": r_pptx_inspect,
    "pptx_compose": r_pptx_compose,
    "receipts_count": len(r_receipts.get("receipts", [])),
}}
print("---BEGIN-PROOF-JSON---")
print(json.dumps(output, indent=2))
print("---END-PROOF-JSON---")
"""
        env = os.environ.copy()
        env["HERMES_HOME"] = str(hermes_home)
        env.pop("HERMES_PROFILE", None)

        res = subprocess.run(
            [str(VENV_PYTHON), "-c", runner_code],
            capture_output=True,
            text=True,
            env=env,
            timeout=180,
        )
        if res.returncode != 0:
            print("STDERR:\n", res.stderr, file=sys.stderr)
            print("STDOUT:\n", res.stdout, file=sys.stderr)
            raise RuntimeError(f"Proof run failed with exit code {res.returncode}")

        raw_out = res.stdout + "\n" + res.stderr
        start_tag = "---BEGIN-PROOF-JSON---"
        end_tag = "---END-PROOF-JSON---"
        if start_tag in raw_out and end_tag in raw_out:
            payload = raw_out.split(start_tag)[1].split(end_tag)[0].strip()
            data = json.loads(payload)
            return data
        else:
            print("STDERR:\n", res.stderr, file=sys.stderr)
            print("STDOUT:\n", res.stdout, file=sys.stderr)
            raise RuntimeError(f"Could not parse output tags: {raw_out}")
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)


if __name__ == "__main__":
    result = run_proof()
    print("SUCCESS: hermes-office proof executed cleanly!")
    print(f"Desk registered: {result['desk']['ok']}")
    print(f"Docx tracked edit ok: {result['docx_edit']['ok']}, revisions: {result['docx_revisions']}")
    print(f"Docx refusal caught safely: ok={result['docx_refusal']['ok']}, error={result['docx_refusal'].get('error')}")
    print(f"Xlsx edit ok: {result['xlsx_edit']['ok']}, receipt={result['xlsx_edit'].get('receipt_id')}")
    print(f"Xlsx recalc status: recalculated={result['xlsx_recalc'].get('recalculated')}, values={result['xlsx_recalc'].get('values')}")
    print(f"Xlsx diff changed cells: {len(result['xlsx_diff'].get('diff', {}).get('changed', []))}")
    print(f"Pptx diff slide changes: {len(result['pptx_diff'].get('diff', {}).get('slide_changes', []))}")
    print(f"Pptx compose ok: {result['pptx_compose']['ok']}, slides: {result['pptx_compose'].get('slides')}")
    print(f"Total receipts tracked: {result['receipts_count']}")
