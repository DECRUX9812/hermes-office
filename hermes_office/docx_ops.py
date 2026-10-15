from __future__ import annotations

from pathlib import Path
from typing import Any

import docx
from docx.commentops import comment_thread, reply, resolve
from docx.errors import AmbiguousTargetError, TargetNotFoundError
from docx.package import compare, patch_save
from docx.search import find_one, find_text

from .guard import atomic_write, backup, kind_of, resolve_path


def _load(path: str | Path) -> docx.Document:
    return docx.Document(str(resolve_path(path)))


def _save(path: str | Path, document: docx.Document) -> Path:
    target = resolve_path(path)
    if target.exists():
        tmp = target.parent / f".{target.name}.patched.docx"
        patch_save(str(target), document, str(tmp))
        tmp.replace(target)
    else:
        document.save(str(target))
    return target


def inspect_docx(path: str | Path, *, detail: str = "summary") -> dict[str, Any]:
    document = _load(path)
    if detail == "full":
        paragraphs = [paragraph.text for paragraph in document.paragraphs]
    else:
        paragraphs = []
    return {
        "kind": "docx",
        "path": str(resolve_path(path)),
        "summary": {
            "paragraphs": len(document.paragraphs),
            "tables": len(document.tables),
            "comments": len(tuple(document.comments)),
            "revisions": len(document.revisions),
        },
        "paragraphs": paragraphs,
    }


def search_docx(path: str | Path, query: str, *, match: str = "exact", regex: bool = False, nth: int | None = None) -> list[dict[str, Any]]:
    document = _load(path)
    if regex:
        spans = find_text(document, query, match="exact")
    elif nth is not None:
        spans = [find_one(document, query, nth=nth, match=match)]
    else:
        spans = find_text(document, query, match=match)
    return [span.to_dict() if hasattr(span, "to_dict") else {"text": getattr(span, "text", query)} for span in spans]


def tracked_edit_docx(
    path: str | Path,
    *,
    find: str,
    replace: str,
    nth: int | None = None,
    match: str = "exact",
    author: str | None = None,
    session_id: str | None = None,
) -> dict[str, Any]:
    document = _load(path)
    span = find_one(document, find, nth=nth, match=match)
    result = span.replace(replace, tracked=True, author=author or "Hermes")
    _save(path, document)
    return {"result": result.to_dict(), "revisions": len(document.revisions)}


def comment_docx(
    path: str | Path,
    *,
    find: str,
    text: str,
    author: str,
    nth: int | None = None,
    reply_to: int | None = None,
    resolve_comment: bool = False,
) -> dict[str, Any]:
    document = _load(path)
    author = author or "Hermes"
    if reply_to is not None:
        thread_comment = next((comment for comment in document.comments if comment.comment_id == reply_to), None)
        if thread_comment is None:
            raise TargetNotFoundError(f"comment {reply_to} not found")
        comment = reply(document, thread_comment, text, author=author)
    else:
        span = find_one(document, find, nth=nth)
        comment = span.comment(text, author=author)
    if resolve_comment:
        resolve(document, comment)
    _save(path, document)
    return {"comment": {"comment_id": comment.comment_id, "author": comment.author, "text": comment.text, "resolved": getattr(comment, "resolved", False)}}


def redline_docx(
    path: str | Path,
    *,
    action: str,
    author: str | None = None,
    compare_with: str | Path | None = None,
    scope: str = "all",
) -> dict[str, Any]:
    document = _load(path)
    if compare_with is not None:
        document = compare(compare_with, path, author=author or "Hermes").document
    if action == "accept":
        resolved = document.revisions.accept_all(author=author or "Hermes")
    elif action == "reject":
        resolved = document.revisions.reject_all(author=author or "Hermes")
    else:
        raise ValueError("action must be accept or reject")
    _save(path, document)
    return {"resolved": resolved, "remaining": len(document.revisions), "scope": scope}


def preview_docx(path: str | Path) -> dict[str, Any]:
    document = _load(path)
    revisions = [revision.to_dict() if hasattr(revision, "to_dict") else {} for revision in document.revisions]
    paragraphs = []
    for index, paragraph in enumerate(document.paragraphs):
        runs = [{"text": run.text, "rev": None, "author": None} for run in paragraph.runs]
        for revision in revisions:
            anchor = revision.get("anchor") or {}
            if anchor.get("index") == index:
                runs.append({"text": revision.get("text", ""), "rev": "ins" if "insert" in revision.get("revision_type", "") else "del", "author": revision.get("author")})
        paragraphs.append({"style": getattr(paragraph.style, "name", None), "runs": runs})
    return {
        "kind": "docx",
        "path": str(resolve_path(path)),
        "docx": {
            "paragraphs": paragraphs,
            "comments": list(comment_thread(document)),
        },
    }


def diff_docx(path: str | Path, baseline: str | Path) -> dict[str, Any]:
    current = _load(path)
    original = _load(baseline)
    before = [p.text for p in original.paragraphs]
    after = [p.text for p in current.paragraphs]
    paragraphs_changed = [
        {"before": before[idx] if idx < len(before) else "", "after": after[idx] if idx < len(after) else ""}
        for idx in range(max(len(before), len(after)))
        if (before[idx] if idx < len(before) else "") != (after[idx] if idx < len(after) else "")
    ]
    revisions = [revision.to_dict() if hasattr(revision, "to_dict") else {} for revision in current.revisions]
    return {"kind": "docx", "revisions": revisions, "paragraphs_changed": paragraphs_changed}


def accept_or_reject_docx(path: str | Path, action: str, *, author: str | None = None, compare_with: str | Path | None = None) -> dict[str, Any]:
    return redline_docx(path, action=action, author=author, compare_with=compare_with)
