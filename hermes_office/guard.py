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

from .errors import OutsideDeskError, UnsupportedFileError

_SUPPORTED = {".docx": "docx", ".pptx": "pptx", ".xlsx": "xlsx"}
_UNSUPPORTED = {".doc", ".ppt", ".xls"}


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
    if suffix == ".xlsm":
        raise UnsupportedFileError("macro-enabled workbooks are not supported")
    if suffix in _SUPPORTED:
        return _SUPPORTED[suffix]
    if suffix in _UNSUPPORTED:
        raise UnsupportedFileError(f"{suffix} files are not supported")
    raise UnsupportedFileError(f"unsupported file type: {suffix or resolved.name}")


def backup(data_dir: str | os.PathLike[str], ws_id: int, receipt_id: int, path: str | os.PathLike[str]) -> Path:
    source = resolve_path(path)
    backup_dir = Path(data_dir) / "backups" / str(ws_id) / str(receipt_id)
    backup_dir.mkdir(parents=True, exist_ok=True)
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


def refusal_payload(exc: Exception) -> dict[str, Any]:
    reason = getattr(exc, "reason", exc.__class__.__name__)
    hint = getattr(exc, "hint", None) or _hint_for_exception(exc)
    if isinstance(exc, (DocxPaperRefusal, PptxPaperRefusal, XlsxPaperRefusal)):
        reason = exc.__class__.__name__
    elif isinstance(exc, (OutsideDeskError, UnsupportedFileError)):
        reason = exc.reason
    return {"ok": False, "refused": True, "reason": reason, "message": str(exc), "hint": hint}


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
