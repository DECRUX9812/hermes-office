from __future__ import annotations

import hashlib
import os
import shutil
import tempfile
from pathlib import Path
from typing import Any

from docx.errors import (
    AmbiguousTargetError as DocxAmbiguousTargetError,
    BoundaryViolationError as DocxBoundaryViolationError,
    DocumentProtectedError,
    MalformedPackageError,
    PaperRefusal as DocxPaperRefusal,
    RelationshipPolicyError as DocxRelationshipPolicyError,
    TargetNotFoundError as DocxTargetNotFoundError,
    UnsupportedStructureError as DocxUnsupportedStructureError,
)
from openpyxl.errors import (
    AmbiguousTargetError as XlsxAmbiguousTargetError,
    BoundaryViolationError as XlsxBoundaryViolationError,
    OracleTimeoutError,
    OracleUnavailableError,
    PaperRefusal as XlsxPaperRefusal,
    RelationshipPolicyError as XlsxRelationshipPolicyError,
    TargetNotFoundError as XlsxTargetNotFoundError,
    UnsupportedStructureError as XlsxUnsupportedStructureError,
)
from pptx.errors import (
    AmbiguousTargetError as PptxAmbiguousTargetError,
    BoundaryViolationError as PptxBoundaryViolationError,
    PaperRefusal as PptxPaperRefusal,
    RelationshipPolicyError as PptxRelationshipPolicyError,
    StaleAnchorError,
    TargetNotFoundError as PptxTargetNotFoundError,
    UnsupportedStructureError as PptxUnsupportedStructureError,
)

from .errors import OfficeError, OutsideDeskError, PolicyRefusal, UnsupportedFileError

_SUPPORTED = {".docx": "docx", ".pptx": "pptx", ".xlsx": "xlsx"}
OFFICE_SUFFIXES = frozenset(_SUPPORTED)


def resolve_path(path: str | os.PathLike[str]) -> Path:
    return Path(path).expanduser().resolve()


def require_on_desk(desk, path: str | os.PathLike[str]) -> dict[str, Any]:
    resolved = resolve_path(path)
    folder = desk.folder_for_path(resolved)
    if folder is None:
        raise OutsideDeskError(str(resolved))
    return folder


def kind_of(path: str | os.PathLike[str]) -> str:
    resolved = resolve_path(path)
    suffix = resolved.suffix.lower()
    if suffix in _SUPPORTED:
        return _SUPPORTED[suffix]
    raise UnsupportedFileError(f"{suffix or resolved.name} files are not supported; use .docx, .pptx or .xlsx (legacy and macro-enabled formats are refused)")


def require_kind(path: str | os.PathLike[str], expected: str) -> None:
    actual = kind_of(path)
    if actual != expected:
        raise UnsupportedFileError(f"this tool edits .{expected} files, got .{actual}")


def backup(data_dir: str | os.PathLike[str], ws_id: int, path: str | os.PathLike[str]) -> Path:
    """Copy `path` into a fresh, never-reused directory under the workstream's backups."""
    source = resolve_path(path)
    root = Path(data_dir) / "backups" / str(ws_id)
    root.mkdir(parents=True, exist_ok=True)
    backup_dir = Path(tempfile.mkdtemp(prefix="op-", dir=str(root)))
    target = backup_dir / source.name
    shutil.copy2(source, target)
    return target


def atomic_write(path: str | os.PathLike[str], writer) -> Path:
    target = resolve_path(path)
    target.parent.mkdir(parents=True, exist_ok=True)
    fd, tmp_name = tempfile.mkstemp(dir=str(target.parent), prefix=f".{target.name}.", suffix=".tmp")
    os.close(fd)
    tmp_path = Path(tmp_name)
    try:
        writer(tmp_path)
        os.replace(tmp_path, target)
    except Exception:
        if tmp_path.exists():
            tmp_path.unlink()
        raise
    return target


def sha256(path: str | os.PathLike[str]) -> str:
    digest = hashlib.sha256()
    with Path(path).open("rb") as fh:
        for chunk in iter(lambda: fh.read(1 << 16), b""):
            digest.update(chunk)
    return digest.hexdigest()


_PAPER_REFUSALS = (DocxPaperRefusal, PptxPaperRefusal, XlsxPaperRefusal)


def is_refusal(exc: BaseException) -> bool:
    return isinstance(exc, _PAPER_REFUSALS + (PolicyRefusal,))


def refusal_payload(exc: Exception) -> dict[str, Any]:
    """Shape any failure as JSON; library and policy refusals are flagged `refused` (file untouched)."""
    if isinstance(exc, _PAPER_REFUSALS):
        return {"ok": False, "refused": True, "reason": exc.__class__.__name__, "message": str(exc), "hint": _hint_for_exception(exc)}
    if isinstance(exc, OfficeError):
        return {"ok": False, "refused": isinstance(exc, PolicyRefusal), "reason": exc.reason, "message": str(exc), "hint": exc.hint}
    if isinstance(exc, FileNotFoundError):
        return {"ok": False, "refused": False, "reason": "not_found", "message": str(exc), "hint": "Check the path; use office_desk list to see desk folders."}
    if isinstance(exc, KeyError):
        return {"ok": False, "refused": False, "reason": "bad_request", "message": f"missing argument {exc}", "hint": "Pass every required argument."}
    if isinstance(exc, (ValueError, TypeError)):
        return {"ok": False, "refused": False, "reason": "bad_request", "message": str(exc), "hint": "Fix the arguments and retry."}
    return {"ok": False, "refused": False, "reason": "error", "message": f"{type(exc).__name__}: {exc}", "hint": "Inspect the file and retry with a narrower request."}


def _hint_for_exception(exc: Exception) -> str:
    if isinstance(exc, (DocxAmbiguousTargetError, PptxAmbiguousTargetError, XlsxAmbiguousTargetError)):
        return "Narrow the target with nth or a more specific query."
    if isinstance(exc, (DocxTargetNotFoundError, PptxTargetNotFoundError, XlsxTargetNotFoundError)):
        return "Check the search term or pick a different cell, shape, or span."
    if isinstance(exc, (DocxUnsupportedStructureError, PptxUnsupportedStructureError, XlsxUnsupportedStructureError)):
        return "Open a simpler Office file."
    if isinstance(exc, (DocxBoundaryViolationError, PptxBoundaryViolationError, XlsxBoundaryViolationError)):
        return "Keep the edit within the document boundaries."
    if isinstance(exc, (DocxRelationshipPolicyError, PptxRelationshipPolicyError, XlsxRelationshipPolicyError)):
        return "Choose a target the file allows changing."
    if isinstance(exc, DocumentProtectedError):
        return "Remove document protection first."
    if isinstance(exc, MalformedPackageError):
        return "Repair the package or use a fresh copy."
    if isinstance(exc, StaleAnchorError):
        return "Refind the shape and retry."
    if isinstance(exc, OracleUnavailableError):
        return "Install LibreOffice or pass a reachable soffice path."
    if isinstance(exc, OracleTimeoutError):
        return "Retry with a longer timeout."
    return "Check the request and try again."
