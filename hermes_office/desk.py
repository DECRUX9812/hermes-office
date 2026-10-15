from __future__ import annotations

import json
import sqlite3
from collections.abc import Iterator
from contextlib import contextmanager
from datetime import datetime, timezone
from pathlib import Path
from typing import Any
from uuid import uuid4


def _utcnow() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="milliseconds")


def _json(value: Any) -> str:
    return json.dumps(value, ensure_ascii=False, sort_keys=True, default=str)


def _dict(row: sqlite3.Row | dict[str, Any] | None) -> dict[str, Any] | None:
    if row is None:
        return None
    if isinstance(row, dict):
        return dict(row)
    return {k: row[k] for k in row.keys()}


class Desk:
    def __init__(self, data_dir: str | Path):
        self.data_dir = Path(data_dir)
        self.data_dir.mkdir(parents=True, exist_ok=True)
        self.db_path = self.data_dir / "desk.db"
        self.backups_dir = self.data_dir / "backups"
        self.backups_dir.mkdir(parents=True, exist_ok=True)
        self._init_db()

    @contextmanager
    def _connect(self) -> Iterator[sqlite3.Connection]:
        # The agent process (tools) and the backend process (plugin_api) share this file.
        conn = sqlite3.connect(self.db_path, timeout=5, isolation_level=None)
        conn.row_factory = sqlite3.Row
        conn.execute("PRAGMA foreign_keys=ON")
        conn.execute("PRAGMA journal_mode=WAL")
        conn.execute("PRAGMA busy_timeout=5000")
        try:
            yield conn
            if conn.in_transaction:
                conn.commit()
        except BaseException:
            if conn.in_transaction:
                conn.rollback()
            raise
        finally:
            conn.close()

    def _init_db(self) -> None:
        with self._connect() as conn:
            conn.executescript(
                """
                CREATE TABLE IF NOT EXISTS folders (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    path TEXT NOT NULL UNIQUE,
                    name TEXT NOT NULL,
                    added_at TEXT NOT NULL
                );
                CREATE TABLE IF NOT EXISTS workstreams (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    folder_id INTEGER REFERENCES folders(id) ON DELETE SET NULL,
                    session_id TEXT,
                    title TEXT NOT NULL,
                    request TEXT,
                    status TEXT NOT NULL,
                    note TEXT,
                    needs_you INTEGER NOT NULL DEFAULT 0,
                    created_at TEXT NOT NULL,
                    updated_at TEXT NOT NULL
                );
                CREATE TABLE IF NOT EXISTS events (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    ts TEXT NOT NULL,
                    workstream_id INTEGER REFERENCES workstreams(id) ON DELETE CASCADE,
                    kind TEXT NOT NULL,
                    tool TEXT,
                    path TEXT,
                    summary TEXT,
                    data TEXT
                );
                CREATE TABLE IF NOT EXISTS files (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    workstream_id INTEGER NOT NULL REFERENCES workstreams(id) ON DELETE CASCADE,
                    path TEXT NOT NULL,
                    kind TEXT NOT NULL,
                    state TEXT NOT NULL,
                    last_receipt_id INTEGER,
                    updated_at TEXT NOT NULL,
                    UNIQUE(workstream_id, path)
                );
                CREATE TABLE IF NOT EXISTS receipts (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    workstream_id INTEGER NOT NULL REFERENCES workstreams(id) ON DELETE CASCADE,
                    tool TEXT NOT NULL,
                    path TEXT NOT NULL,
                    op_summary TEXT NOT NULL,
                    before_sha256 TEXT,
                    after_sha256 TEXT,
                    backup_path TEXT,
                    library_report TEXT NOT NULL,
                    created_at TEXT NOT NULL,
                    reverted_at TEXT
                );
                """
            )

    def add_folder(self, path: str | Path) -> dict[str, Any]:
        resolved = Path(path).expanduser().resolve()
        with self._connect() as conn:
            conn.execute("BEGIN IMMEDIATE")
            conn.execute(
                "INSERT INTO folders(path, name, added_at) VALUES (?, ?, ?) ON CONFLICT(path) DO NOTHING",
                (str(resolved), resolved.name, _utcnow()),
            )
            row = conn.execute("SELECT * FROM folders WHERE path = ?", (str(resolved),)).fetchone()
        return _folder_row(self, row)

    def remove_folder(self, folder_id: int) -> None:
        with self._connect() as conn:
            conn.execute("BEGIN IMMEDIATE")
            conn.execute("DELETE FROM folders WHERE id = ?", (folder_id,))

    def get_folder(self, folder_id: int) -> dict[str, Any] | None:
        with self._connect() as conn:
            row = conn.execute("SELECT * FROM folders WHERE id = ?", (folder_id,)).fetchone()
        return _folder_row(self, row) if row is not None else None

    def list_folders(self) -> list[dict[str, Any]]:
        with self._connect() as conn:
            rows = conn.execute("SELECT * FROM folders ORDER BY added_at DESC, id DESC").fetchall()
        return [_folder_row(self, row) for row in rows]

    def folder_for_path(self, path: str | Path) -> dict[str, Any] | None:
        resolved = Path(path).expanduser().resolve()
        folders = self.list_folders()
        matches = []
        for folder in folders:
            folder_path = Path(folder["path"])
            if resolved == folder_path or folder_path in resolved.parents:
                matches.append(folder)
        if not matches:
            return None
        return max(matches, key=lambda folder: len(Path(folder["path"]).parts))

    def create_workstream(
        self,
        folder_id: int | None,
        title: str,
        request: str | None = None,
        session_id: str | None = None,
        status: str = "queued",
    ) -> dict[str, Any]:
        now = _utcnow()
        with self._connect() as conn:
            conn.execute("BEGIN IMMEDIATE")
            conn.execute(
                """
                INSERT INTO workstreams(folder_id, session_id, title, request, status, note, needs_you, created_at, updated_at)
                VALUES (?, ?, ?, ?, ?, ?, 0, ?, ?)
                """,
                (folder_id, session_id, title, request, status, None, now, now),
            )
            ws_id = conn.execute("SELECT last_insert_rowid()").fetchone()[0]
        return self.get_workstream(ws_id)

    def workstream_for_session(self, session_id: str) -> dict[str, Any] | None:
        with self._connect() as conn:
            row = conn.execute(
                "SELECT * FROM workstreams WHERE session_id = ? ORDER BY updated_at DESC, id DESC LIMIT 1",
                (session_id,),
            ).fetchone()
        return _workstream_row(self, row)

    def get_workstream(self, workstream_id: int) -> dict[str, Any] | None:
        with self._connect() as conn:
            row = conn.execute("SELECT * FROM workstreams WHERE id = ?", (workstream_id,)).fetchone()
        return _workstream_row(self, row)

    def update_workstream(self, workstream_id: int, **fields: Any) -> dict[str, Any]:
        allowed = {k: fields[k] for k in ("folder_id", "session_id", "title", "request", "status", "note", "needs_you") if k in fields}
        if not allowed:
            return self.get_workstream(workstream_id)
        allowed["updated_at"] = _utcnow()
        columns = ", ".join(f"{key} = ?" for key in allowed)
        with self._connect() as conn:
            conn.execute("BEGIN IMMEDIATE")
            conn.execute(f"UPDATE workstreams SET {columns} WHERE id = ?", (*allowed.values(), workstream_id))
        return self.get_workstream(workstream_id)

    def list_workstreams(self, folder_id: int | None = None, status: str | None = None, limit: int = 200) -> list[dict[str, Any]]:
        query = "SELECT * FROM workstreams"
        clauses = []
        params: list[Any] = []
        if folder_id is not None:
            clauses.append("folder_id = ?")
            params.append(folder_id)
        if status is not None:
            clauses.append("status = ?")
            params.append(status)
        if clauses:
            query += " WHERE " + " AND ".join(clauses)
        query += " ORDER BY updated_at DESC, id DESC LIMIT ?"
        params.append(limit)
        with self._connect() as conn:
            rows = conn.execute(query, params).fetchall()
        return [_workstream_row(self, row) for row in rows]

    def add_event(
        self,
        workstream_id: int,
        kind: str,
        tool: str | None = None,
        path: str | None = None,
        summary: str | None = None,
        data: dict[str, Any] | None = None,
    ) -> dict[str, Any]:
        payload = dict(data or {})
        payload.setdefault("op_id", str(uuid4()))
        ts = _utcnow()
        with self._connect() as conn:
            conn.execute("BEGIN IMMEDIATE")
            conn.execute(
                "INSERT INTO events(ts, workstream_id, kind, tool, path, summary, data) VALUES (?, ?, ?, ?, ?, ?, ?)",
                (ts, workstream_id, kind, tool, path, summary, _json(payload)),
            )
            row = conn.execute("SELECT * FROM events WHERE id = last_insert_rowid()").fetchone()
        return _event_row(row) or {}

    def now(self, workstream_id: int) -> dict[str, Any] | None:
        events = self.events(workstream_id)
        seen_done: set[str] = set()
        for event in reversed(events):
            if event["kind"] in {"tool_finished", "refused"}:
                seen_done.add((event.get("data") or {}).get("op_id", ""))
        for event in reversed(events):
            if event["kind"] == "tool_started":
                op_id = (event.get("data") or {}).get("op_id", "")
                if op_id not in seen_done:
                    return {
                        "label": event.get("summary") or event.get("tool") or "",
                        "tool": event.get("tool"),
                        "path": event.get("path"),
                        "since": event.get("ts"),
                    }
        return None

    def track_file(
        self,
        workstream_id: int,
        path: str | Path,
        state: str,
        receipt_id: int | None = None,
        kind: str | None = None,
    ) -> dict[str, Any]:
        resolved = str(Path(path).expanduser().resolve())
        now = _utcnow()
        if kind is None:
            kind = "unknown"
        with self._connect() as conn:
            conn.execute("BEGIN IMMEDIATE")
            conn.execute(
                """
                INSERT INTO files(workstream_id, path, kind, state, last_receipt_id, updated_at)
                VALUES (?, ?, ?, ?, ?, ?)
                ON CONFLICT(workstream_id, path) DO UPDATE SET
                    kind=excluded.kind,
                    state=excluded.state,
                    last_receipt_id=excluded.last_receipt_id,
                    updated_at=excluded.updated_at
                """,
                (workstream_id, resolved, kind, state, receipt_id, now),
            )
            row = conn.execute(
                "SELECT * FROM files WHERE workstream_id = ? AND path = ?", (workstream_id, resolved)
            ).fetchone()
        return _dict(row) or {}

    def files(self, workstream_id: int) -> list[dict[str, Any]]:
        with self._connect() as conn:
            rows = conn.execute("SELECT * FROM files WHERE workstream_id = ? ORDER BY updated_at DESC, id DESC", (workstream_id,)).fetchall()
        return [_dict(row) or {} for row in rows]

    def add_receipt(
        self,
        workstream_id: int,
        tool: str,
        path: str | Path,
        op_summary: str,
        before_sha256: str | None,
        after_sha256: str | None,
        backup_path: str | Path | None,
        library_report: dict[str, Any],
    ) -> dict[str, Any]:
        ts = _utcnow()
        with self._connect() as conn:
            conn.execute("BEGIN IMMEDIATE")
            conn.execute(
                """
                INSERT INTO receipts(workstream_id, tool, path, op_summary, before_sha256, after_sha256, backup_path, library_report, created_at, reverted_at)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, NULL)
                """,
                (
                    workstream_id,
                    tool,
                    str(Path(path).expanduser().resolve()),
                    op_summary,
                    before_sha256,
                    after_sha256,
                    str(backup_path) if backup_path is not None else None,
                    _json(library_report),
                    ts,
                ),
            )
            receipt_id = conn.execute("SELECT last_insert_rowid()").fetchone()[0]
        return self.get_receipt(receipt_id) or {}

    def get_receipt(self, receipt_id: int) -> dict[str, Any] | None:
        with self._connect() as conn:
            row = conn.execute("SELECT * FROM receipts WHERE id = ?", (receipt_id,)).fetchone()
        return _receipt_row(row)

    def list_receipts(self, workstream_id: int | None = None) -> list[dict[str, Any]]:
        query = "SELECT * FROM receipts"
        params: list[Any] = []
        if workstream_id is not None:
            query += " WHERE workstream_id = ?"
            params.append(workstream_id)
        query += " ORDER BY created_at DESC, id DESC"
        with self._connect() as conn:
            rows = conn.execute(query, params).fetchall()
        return [_receipt_row(row) or {} for row in rows]

    def mark_reverted(self, receipt_id: int) -> dict[str, Any] | None:
        ts = _utcnow()
        with self._connect() as conn:
            conn.execute("BEGIN IMMEDIATE")
            conn.execute("UPDATE receipts SET reverted_at = ? WHERE id = ?", (ts, receipt_id))
        return self.get_receipt(receipt_id)

    def events_since(self, ts: str | None) -> list[dict[str, Any]]:
        query = "SELECT * FROM events"
        params: list[Any] = []
        if ts:
            query += " WHERE ts > ?"
            params.append(ts)
        query += " ORDER BY id ASC LIMIT 500"
        with self._connect() as conn:
            rows = conn.execute(query, params).fetchall()
        return [_event_row(row) or {} for row in rows]

    def first_receipt_for(self, path: str | Path, workstream_id: int | None = None) -> dict[str, Any] | None:
        query = "SELECT * FROM receipts WHERE path = ? AND backup_path IS NOT NULL"
        params: list[Any] = [str(Path(path).expanduser().resolve())]
        if workstream_id is not None:
            query += " AND workstream_id = ?"
            params.append(workstream_id)
        with self._connect() as conn:
            row = conn.execute(query + " ORDER BY id ASC LIMIT 1", params).fetchone()
        return _receipt_row(row)

    def events(self, workstream_id: int) -> list[dict[str, Any]]:
        with self._connect() as conn:
            rows = conn.execute("SELECT * FROM events WHERE workstream_id = ? ORDER BY id ASC", (workstream_id,)).fetchall()
        return [_event_row(row) or {} for row in rows]


def _folder_row(desk: Desk, row: sqlite3.Row | None) -> dict[str, Any]:
    data = _dict(row) or {}
    if not data:
        return data
    folder_id = data["id"]
    with desk._connect() as conn:
        data["counts"] = {
            "files": conn.execute("SELECT COUNT(DISTINCT path) FROM files WHERE workstream_id IN (SELECT id FROM workstreams WHERE folder_id = ?)", (folder_id,)).fetchone()[0],
            "workstreams": conn.execute("SELECT COUNT(*) FROM workstreams WHERE folder_id = ?", (folder_id,)).fetchone()[0],
            "needs_you": conn.execute("SELECT COUNT(*) FROM workstreams WHERE folder_id = ? AND needs_you = 1", (folder_id,)).fetchone()[0],
        }
    return data


def _workstream_row(desk: Desk, row: sqlite3.Row | None) -> dict[str, Any]:
    data = _dict(row) or {}
    if not data:
        return data
    workstream_id = data["id"]
    with desk._connect() as conn:
        last_event = conn.execute(
            "SELECT * FROM events WHERE workstream_id = ? ORDER BY id DESC LIMIT 1", (workstream_id,)
        ).fetchone()
        data["files_changed"] = conn.execute(
            "SELECT COUNT(*) FROM files WHERE workstream_id = ? AND state IN ('changed', 'created')", (workstream_id,)
        ).fetchone()[0]
    data["last_event"] = _event_row(last_event)
    return data


def _receipt_row(row: sqlite3.Row | None) -> dict[str, Any] | None:
    data = _dict(row)
    if data is None:
        return None
    if data.get("library_report"):
        try:
            data["library_report"] = json.loads(data["library_report"])
        except Exception:
            pass
    return data


def _event_row(row: sqlite3.Row | None | dict[str, Any]) -> dict[str, Any] | None:
    data = _dict(row)
    if data is None:
        return None
    if data.get("data"):
        try:
            data["data"] = json.loads(data["data"])
        except Exception:
            pass
    return data
