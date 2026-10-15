from __future__ import annotations

from typing import Any

_WS = {"type": "integer", "description": "Desk workstream id, when the request named one. Usually omit: the current chat's workstream is used."}
_PATH = {"type": "string", "description": "Absolute path to the file."}


def _schema(name: str, description: str, properties: dict[str, Any], required: list[str]) -> dict[str, Any]:
    return {"name": name, "description": description, "parameters": {"type": "object", "properties": properties, "required": required}}


OFFICE_DESK_SCHEMA = _schema(
    "office_desk",
    "The user's Desk: folders the agent may edit Office files in, and workstreams (one per request) the user reviews. "
    "list = folders + recent workstreams. add_folder only when the user asked for that folder. "
    "needs_you = hand the workstream to the user with a note (a question, or 'please review'). done = mark finished.",
    {
        "action": {"type": "string", "enum": ["list", "add_folder", "remove_folder", "status", "needs_you", "done", "attach"]},
        "path": {"type": "string", "description": "add_folder: absolute folder path."},
        "folder_id": {"type": "integer", "description": "remove_folder: folder id."},
        "workstream_id": _WS,
        "note": {"type": "string", "description": "needs_you/done: one or two sentences for the user."},
        "title": {"type": "string", "description": "attach: optional new title."},
    },
    ["action"],
)

OFFICE_INSPECT_SCHEMA = _schema(
    "office_inspect",
    "Read a .docx/.pptx/.xlsx file's structure before editing: paragraphs, headings, tables, comments and tracked changes "
    "(Word); slides, shape names and text (PowerPoint); sheets, dimensions, formulas and error cells (Excel). Read-only.",
    {"path": _PATH, "detail": {"type": "string", "enum": ["summary", "full"], "description": "summary (default) or full."}, "workstream_id": _WS},
    ["path"],
)

OFFICE_SEARCH_SCHEMA = _schema(
    "office_search",
    "Find text in one Office file (path) or every Office file under a folder. Use it to pick an exact, unique target "
    "before a tracked edit. Read-only.",
    {
        "path": _PATH,
        "folder": {"type": "string", "description": "Search every .docx/.pptx/.xlsx under this folder instead of one file."},
        "query": {"type": "string"},
        "match": {"type": "string", "enum": ["exact", "normalized"], "description": "normalized folds case, whitespace and quotes."},
        "regex": {"type": "boolean", "description": "Treat query as a Python regular expression."},
    },
    ["query"],
)

DOCX_TRACKED_EDIT_SCHEMA = _schema(
    "docx_tracked_edit",
    "Replace text in a Word document as a TRACKED CHANGE (w:del + w:ins, visible in Word's Review pane) so the user can "
    "accept or reject it. The find text must match exactly one place unless nth is given; ambiguous or missing targets are "
    "refused and the file is left unchanged. Backs the file up first and records a receipt.",
    {
        "path": _PATH,
        "find": {"type": "string", "description": "Exact text currently in the document."},
        "replace": {"type": "string", "description": "New text."},
        "nth": {"type": "integer", "description": "0-based occurrence when find appears more than once."},
        "match": {"type": "string", "enum": ["exact", "normalized"]},
        "author": {"type": "string", "description": "Revision author shown in Word (default: the plugin setting Tracked-change author, Hermes unless changed)."},
        "workstream_id": _WS,
    },
    ["path", "find", "replace"],
)

DOCX_COMMENT_SCHEMA = _schema(
    "docx_comment",
    "Add a Word comment anchored on exact text (find), reply to a comment (reply_to), or resolve one (resolve). "
    "Comment ids come from office_inspect detail=full.",
    {
        "path": _PATH,
        "text": {"type": "string", "description": "Comment or reply text."},
        "find": {"type": "string", "description": "Exact text to anchor a new comment on."},
        "nth": {"type": "integer", "description": "0-based occurrence of find."},
        "reply_to": {"type": "integer", "description": "Comment id to reply to."},
        "resolve": {"type": "integer", "description": "Comment id to mark resolved."},
        "author": {"type": "string"},
        "workstream_id": _WS,
    },
    ["path"],
)

DOCX_REDLINE_SCHEMA = _schema(
    "docx_redline_accept_reject",
    "Resolve tracked changes in a Word document (accept or reject: all, one author's, or one by index), or build a redline "
    "(compare: tracked changes that turn compare_with, the original, into path, the revised version; written to output_path). "
    "Only accept/reject when the user asked to.",
    {
        "path": _PATH,
        "action": {"type": "string", "enum": ["accept", "reject", "compare"]},
        "scope": {"type": "string", "enum": ["all", "author", "index"], "description": "accept/reject selection (default all)."},
        "author": {"type": "string", "description": "scope=author: whose changes; compare: redline author."},
        "index": {"type": "integer", "description": "scope=index: 0-based revision index from office_inspect detail=full."},
        "compare_with": {"type": "string", "description": "compare: absolute path of the ORIGINAL document."},
        "output_path": {"type": "string", "description": "compare: where to write the redline (default '<name> (redline).docx' beside path)."},
        "workstream_id": _WS,
    },
    ["path", "action"],
)

PPTX_INSPECT_FORMATS_SCHEMA = _schema(
    "pptx_inspect_formats",
    "Effective formatting of one named shape on a slide (font, size, colour, paragraph and shape format) with provenance: "
    "which layer (run, placeholder, layout, master, theme) supplied each value. Read-only.",
    {"path": _PATH, "slide": {"type": "integer", "description": "0-based slide index."}, "shape": {"type": "string", "description": "Shape name from office_inspect."}},
    ["path", "slide", "shape"],
)

PPTX_EDIT_SHAPE_SCHEMA = _schema(
    "pptx_edit_shape",
    "Edit a PowerPoint deck in place, keeping formatting. Ops run in order on one slide: replace_text (in the named shape), "
    "replace_text_deck (every slide), set_geometry (EMU x/y/w/h of the named shape), delete_shape, clone_slide (inserted after), "
    "delete_slide, move_slide (to_index). Missing or duplicate shape names are refused.",
    {
        "path": _PATH,
        "slide": {"type": "integer", "description": "0-based slide index."},
        "shape": {"type": "string", "description": "Shape name for shape ops."},
        "ops": {
            "type": "array",
            "items": {
                "type": "object",
                "properties": {
                    "op": {"type": "string", "enum": ["replace_text", "replace_text_deck", "set_geometry", "delete_shape", "clone_slide", "delete_slide", "move_slide"]},
                    "find": {"type": "string"},
                    "replace": {"type": "string"},
                    "include_notes": {"type": "boolean"},
                    "x": {"type": "integer"}, "y": {"type": "integer"}, "w": {"type": "integer"}, "h": {"type": "integer"},
                    "to_index": {"type": "integer"},
                },
                "required": ["op"],
            },
        },
        "workstream_id": _WS,
    },
    ["path", "slide", "ops"],
)

PPTX_COMPOSE_SCHEMA = _schema(
    "pptx_compose_decks",
    "Import slides from other decks into a deck. mode is required: adopt_theme (take the target deck's look), "
    "keep_appearance (carry the source layout/master), bake (flatten the source look onto the slide). "
    "Writes output_path (default: the target deck in place).",
    {
        "path": {"type": "string", "description": "Target deck (absolute path)."},
        "sources": {
            "type": "array",
            "items": {
                "type": "object",
                "properties": {"path": {"type": "string"}, "slides": {"type": "array", "items": {"type": "integer"}, "description": "0-based indexes; omit for all."}},
                "required": ["path"],
            },
        },
        "mode": {"type": "string", "enum": ["adopt_theme", "keep_appearance", "bake"]},
        "output_path": {"type": "string"},
        "workstream_id": _WS,
    },
    ["path", "sources", "mode"],
)

PPTX_DIFF_SCHEMA = _schema(
    "pptx_diff",
    "Compare two decks: slides added/removed/moved and per-shape text changes. Read-only.",
    {"a": {"type": "string", "description": "Before deck."}, "b": {"type": "string", "description": "After deck."}, "detail": {"type": "string", "enum": ["structure", "text", "full"]}},
    ["a", "b"],
)

XLSX_EDIT_SCHEMA = _schema(
    "xlsx_edit",
    "Edit a workbook in place, preserving everything you do not touch (styles, charts, pivots, defined names). Ops: set "
    "(value), formula, insert_rows/delete_rows/insert_cols/delete_cols (formulas are rewritten reference-aware), rename_sheet. "
    "Returns the library's edit receipt (cells and parts changed). Run xlsx_recalc afterwards to refresh cached values.",
    {
        "path": _PATH,
        "ops": {
            "type": "array",
            "items": {
                "type": "object",
                "properties": {
                    "op": {"type": "string", "enum": ["set", "formula", "insert_rows", "delete_rows", "insert_cols", "delete_cols", "rename_sheet"]},
                    "ref": {"type": "string", "description": "Cell like B3 or Sheet!B3."},
                    "value": {"type": "string", "description": "set: text value; formula: '=SUM(B2:B4)'."},
                    "number": {"type": "number", "description": "set: numeric value (use instead of value for numbers)."},
                    "sheet": {"type": "string"},
                    "index": {"type": "integer", "description": "1-based row/column index."},
                    "amount": {"type": "integer"},
                    "new_name": {"type": "string"},
                },
                "required": ["op"],
            },
        },
        "workstream_id": _WS,
    },
    ["path", "ops"],
)

XLSX_RECALC_SCHEMA = _schema(
    "xlsx_recalc",
    "Recalculate a workbook's formulas with LibreOffice and write the fresh cached values back (preserving the package). "
    "Returns the values of the cells in read and any formula errors. Without LibreOffice it says so and changes nothing.",
    {"path": _PATH, "read": {"type": "array", "items": {"type": "string"}, "description": "Cells to report, like Sheet!B8."}, "workstream_id": _WS},
    ["path"],
)

XLSX_DIFF_SCHEMA = _schema(
    "xlsx_diff",
    "Cell-level diff of two workbooks (changed values/formulas, shifted cells, added/removed sheets). Read-only.",
    {"a": {"type": "string", "description": "Before workbook."}, "b": {"type": "string", "description": "After workbook."}},
    ["a", "b"],
)

OFFICE_RECEIPT_SCHEMA = _schema(
    "office_receipt",
    "Receipts are the record of every file change (hashes, backup, library report). list (optionally per workstream), "
    "show one, or revert one (restores the backup; itself receipted).",
    {"action": {"type": "string", "enum": ["list", "show", "revert"]}, "workstream_id": _WS, "receipt_id": {"type": "integer"}},
    ["action"],
)

ALL_SCHEMAS = {
    s["name"]: s
    for s in (
        OFFICE_DESK_SCHEMA, OFFICE_INSPECT_SCHEMA, OFFICE_SEARCH_SCHEMA, DOCX_TRACKED_EDIT_SCHEMA, DOCX_COMMENT_SCHEMA,
        DOCX_REDLINE_SCHEMA, PPTX_INSPECT_FORMATS_SCHEMA, PPTX_EDIT_SHAPE_SCHEMA, PPTX_COMPOSE_SCHEMA, PPTX_DIFF_SCHEMA,
        XLSX_EDIT_SCHEMA, XLSX_RECALC_SCHEMA, XLSX_DIFF_SCHEMA, OFFICE_RECEIPT_SCHEMA,
    )
}
