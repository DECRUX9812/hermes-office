#!/usr/bin/env python3
"""Dev harness server for the Desk desktop page.

Serves desktop/dev/ static files at / and mounts the REAL backend router from
<repo>/dashboard/plugin_api.py at /api/plugins/hermes-office, so the built
plugin.js renders in a plain browser against real data. HERMES_HOME comes from
the environment (the plugin's data dir derives from it).

Run with the office venv python:

    HERMES_HOME=/tmp/desk-hermes-home ~/Code/ship-test/office-venv/.venv/bin/python \
        desktop/dev/serve.py --repo ~/Code/hermes-office-wt/e1 --port 9341
"""

from __future__ import annotations

import argparse
import importlib.util
import os
import sys
from pathlib import Path

import uvicorn
from fastapi import FastAPI
from fastapi.staticfiles import StaticFiles

DEV_DIR = Path(__file__).resolve().parent


def load_router(repo: Path):
    api_path = repo / "dashboard" / "plugin_api.py"
    if not api_path.is_file():
        raise SystemExit(f"no plugin_api.py at {api_path}")

    # The router imports `hermes_office` from the repo; make sure the repo wins
    # over whatever else is importable in this interpreter.
    sys.path.insert(0, str(repo))

    spec = importlib.util.spec_from_file_location("hermes_office_plugin_api", api_path)
    if spec is None or spec.loader is None:
        raise SystemExit(f"could not load {api_path}")

    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)

    router = getattr(module, "router", None)
    if router is None:
        raise SystemExit(f"{api_path} does not export `router`")

    return router


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--repo", required=True, type=Path, help="path to the hermes-office plugin repo")
    parser.add_argument("--port", type=int, default=9341)
    parser.add_argument("--host", default="127.0.0.1")
    args = parser.parse_args()

    if not os.environ.get("HERMES_HOME"):
        raise SystemExit("set HERMES_HOME (a temp dir, never ~/.hermes) before running the harness")

    app = FastAPI(title="hermes-office desk dev harness")
    app.include_router(load_router(args.repo.resolve()), prefix="/api/plugins/hermes-office")

    dist = DEV_DIR / "dist"
    if not (dist / "harness.js").is_file():
        raise SystemExit("build the harness first: node desktop/dev/build-harness.mjs")

    # html=True serves index.html at /; static mounts after the API router.
    app.mount("/", StaticFiles(directory=DEV_DIR, html=True), name="harness")

    uvicorn.run(app, host=args.host, port=args.port, log_level="warning")


if __name__ == "__main__":
    main()
