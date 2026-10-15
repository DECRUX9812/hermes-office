from __future__ import annotations

from pathlib import Path
from typing import Any

from .docx_ops import preview_docx
from .guard import kind_of
from .pptx_ops import preview_pptx
from .xlsx_ops import preview_xlsx


def preview_file(path: str | Path) -> dict[str, Any]:
    kind = kind_of(path)
    if kind == "docx":
        return preview_docx(path)
    if kind == "pptx":
        return preview_pptx(path)
    if kind == "xlsx":
        return preview_xlsx(path)
    raise ValueError(f"unsupported kind: {kind}")
