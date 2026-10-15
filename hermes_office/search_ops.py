from __future__ import annotations

from pathlib import Path
from typing import Any

from .docx_ops import search_docx
from .guard import OFFICE_SUFFIXES, kind_of, resolve_path
from .pptx_ops import search_pptx
from .xlsx_ops import search_xlsx

MAX_FILES = 200
MAX_DEPTH = 4


def _search_one(path: Path, query: str, *, match: str, regex: bool) -> list[dict[str, Any]]:
    kind = kind_of(path)
    if kind == "docx":
        return search_docx(path, query, match=match, regex=regex)
    if kind == "pptx":
        return search_pptx(path, query, match=match, regex=regex)
    return search_xlsx(path, query, regex=regex)


def office_files(folder: Path, *, max_depth: int = MAX_DEPTH, limit: int = MAX_FILES) -> list[Path]:
    found: list[Path] = []
    for child in sorted(folder.rglob("*")):
        rel = child.relative_to(folder).parts
        if len(rel) > max_depth or any(part.startswith(".") for part in rel) or child.name.startswith("~$"):
            continue
        if child.is_file() and child.suffix.lower() in OFFICE_SUFFIXES:
            found.append(child)
            if len(found) >= limit:
                break
    return found


def search_path_or_folder(path: str | None, folder: str | None, *, query: str, match: str = "exact", regex: bool = False) -> list[dict[str, Any]]:
    target = folder or path
    if not target:
        raise ValueError("give path (one file) or folder")
    resolved = resolve_path(target)
    files = office_files(resolved) if resolved.is_dir() else [resolved]
    results = []
    for file in files:
        try:
            matches = _search_one(file, query, match=match, regex=regex)
        except Exception as exc:  # one unreadable file must not hide hits in the others
            results.append({"path": str(file), "error": f"{type(exc).__name__}: {exc}"})
            continue
        if matches or not resolved.is_dir():
            results.append({"path": str(file), "matches": matches})
    return results
