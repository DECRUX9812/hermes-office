from __future__ import annotations

from typing import Any


def _schema(name: str, description: str, properties: dict[str, Any], required: list[str] | None = None) -> dict[str, Any]:
    schema = {
        "name": name,
        "description": description,
        "parameters": {
            "type": "object",
            "properties": properties,
            "additionalProperties": False,
        },
    }
    if required:
        schema["parameters"]["required"] = required
    return schema


OFFICE_DESK_SCHEMA = _schema(
    "office_desk",
    "Manage desk folders and workstreams for office files.",
    {
        "action": {"type": "string", "enum": ["add_folder", "remove_folder", "list", "status", "needs_you", "done", "attach"]},
        "path": {"type": "string"},
        "folder_id": {"type": "integer"},
        "workstream_id": {"type": "integer"},
        "session_id": {"type": "string"},
        "title": {"type": "string"},
        "note": {"type": "string"},
    },
    ["action"],
)

OFFICE_INSPECT_SCHEMA = _schema(
    "office_inspect",
    "Inspect an office file and summarize its structure.",
    {"path": {"type": "string"}, "detail": {"type": "string", "enum": ["summary", "full"]}},
    ["path"],
)

OFFICE_SEARCH_SCHEMA = _schema(
    "office_search",
    "Search office text in a file or folder.",
    {
        "path": {"type": "string"},
        "folder": {"type": "string"},
        "query": {"type": "string"},
        "match": {"type": "string", "enum": ["exact", "normalized"]},
        "regex": {"type": "boolean"},
    },
    ["query"],
)

DOCX_TRACKED_EDIT_SCHEMA = _schema(
    "docx_tracked_edit",
    "Make a tracked replacement in a Word document.",
    {
        "path": {"type": "string"},
        "find": {"type": "string"},
        "replace": {"type": "string"},
        "nth": {"type": "integer"},
        "match": {"type": "string", "enum": ["exact", "normalized"]},
        "author": {"type": "string"},
        "workstream_id": {"type": "integer"},
    },
    ["path", "find", "replace"],
)

DOCX_COMMENT_SCHEMA = _schema(
    "docx_comment",
    "Add, reply to, or resolve a Word comment.",
    {
        "path": {"type": "string"},
        "find": {"type": "string"},
        "text": {"type": "string"},
        "author": {"type": "string"},
        "nth": {"type": "integer"},
        "reply_to": {"type": "integer"},
        "resolve": {"type": "boolean"},
        "workstream_id": {"type": "integer"},
    },
    ["path", "find", "text", "author"],
)

DOCX_REDLINE_SCHEMA = _schema(
    "docx_redline_accept_reject",
    "Accept or reject tracked Word changes.",
    {
        "path": {"type": "string"},
        "action": {"type": "string", "enum": ["accept", "reject"]},
        "scope": {"type": "string"},
        "author": {"type": "string"},
        "compare_with": {"type": "string"},
        "workstream_id": {"type": "integer"},
    },
    ["path", "action"],
)

PPTX_INSPECT_FORMATS_SCHEMA = _schema(
    "pptx_inspect_formats",
    "Inspect the style of a named PowerPoint shape.",
    {"path": {"type": "string"}, "slide": {"type": "integer"}, "shape": {"type": "string"}},
    ["path", "slide", "shape"],
)

PPTX_EDIT_SHAPE_SCHEMA = _schema(
    "pptx_edit_shape",
    "Edit or remove a named PowerPoint shape.",
    {
        "path": {"type": "string"},
        "slide": {"type": "integer"},
        "shape": {"type": "string"},
        "ops": {"type": "array"},
        "workstream_id": {"type": "integer"},
    },
    ["path", "slide", "shape", "ops"],
)

PPTX_COMPOSE_SCHEMA = _schema(
    "pptx_compose_decks",
    "Compose a deck from slides in source decks.",
    {
        "dest": {"type": "string"},
        "sources": {"type": "array"},
        "mode": {"type": "string", "enum": ["adopt_theme", "keep_appearance", "bake"]},
        "output_path": {"type": "string"},
        "workstream_id": {"type": "integer"},
    },
    ["dest", "sources", "mode", "output_path"],
)

PPTX_DIFF_SCHEMA = _schema(
    "pptx_diff",
    "Diff two PowerPoint decks.",
    {"a": {"type": "string"}, "b": {"type": "string"}, "detail": {"type": "string"}},
    ["a", "b"],
)

XLSX_EDIT_SCHEMA = _schema(
    "xlsx_edit",
    "Edit a workbook by setting cells or changing structure.",
    {"path": {"type": "string"}, "ops": {"type": "array"}, "workstream_id": {"type": "integer"}},
    ["path", "ops"],
)

XLSX_RECALC_SCHEMA = _schema(
    "xlsx_recalc",
    "Recalculate formulas in a workbook.",
    {"path": {"type": "string"}, "read": {"type": "array"}, "workstream_id": {"type": "integer"}},
    ["path", "read"],
)

XLSX_DIFF_SCHEMA = _schema(
    "xlsx_diff",
    "Diff two workbooks.",
    {"a": {"type": "string"}, "b": {"type": "string"}},
    ["a", "b"],
)

OFFICE_RECEIPT_SCHEMA = _schema(
    "office_receipt",
    "List, show, or revert desk receipts.",
    {
        "action": {"type": "string", "enum": ["list", "show", "revert"]},
        "workstream_id": {"type": "integer"},
        "receipt_id": {"type": "integer"},
    },
    ["action"],
)
