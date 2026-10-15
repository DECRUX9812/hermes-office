from __future__ import annotations

from pathlib import Path
from typing import Any

from .docx_ops import inspect_docx
from .guard import kind_of, resolve_path
from .pptx_ops import inspect_pptx
from .xlsx_ops import inspect_xlsx


def inspect_file(path: str | Path, *, detail: str = "summary") -> dict[str, Any]:
    kind = kind_of(path)
    if kind == "docx":
        return inspect_docx(path, detail=detail)
    if kind == "pptx":
        return inspect_pptx(path)
    if kind == "xlsx":
        return inspect_xlsx(path, detail=detail)
    raise ValueError(f"unsupported kind: {kind}")
