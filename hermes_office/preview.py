from __future__ import annotations

from pathlib import Path
from typing import Any

from .docx_ops import preview_docx
from .guard import kind_of
from .pptx_ops import preview_pptx
from .xlsx_ops import preview_xlsx

_PREVIEWERS = {"docx": preview_docx, "pptx": preview_pptx, "xlsx": preview_xlsx}


def preview_file(path: str | Path) -> dict[str, Any]:
    return _PREVIEWERS[kind_of(path)](path)
