from __future__ import annotations

from pathlib import Path
from typing import Any

from .docx_ops import inspect_docx
from .guard import kind_of
from .pptx_ops import inspect_pptx
from .xlsx_ops import inspect_xlsx

_INSPECTORS = {"docx": inspect_docx, "pptx": inspect_pptx, "xlsx": inspect_xlsx}


def inspect_file(path: str | Path, *, detail: str = "summary") -> dict[str, Any]:
    return _INSPECTORS[kind_of(path)](path, detail=detail)
