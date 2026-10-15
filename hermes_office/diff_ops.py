from __future__ import annotations

from pathlib import Path
from typing import Any

from .docx_ops import diff_docx
from .guard import kind_of, resolve_path
from .pptx_ops import diff_pptx
from .xlsx_ops import diff_xlsx


def diff_file(path: str | Path, baseline: str | Path, *, detail: str = "text") -> dict[str, Any]:
    kind = kind_of(path)
    if kind == "docx":
        return diff_docx(path, baseline)
    if kind == "pptx":
        return diff_pptx(path, baseline, detail=detail)
    if kind == "xlsx":
        return diff_xlsx(path, baseline)
    raise ValueError(f"unsupported kind: {kind}")
