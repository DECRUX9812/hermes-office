from __future__ import annotations

from pathlib import Path
from typing import Any

from .docx_ops import diff_docx
from .guard import kind_of
from .pptx_ops import diff_pptx
from .xlsx_ops import diff_xlsx


def diff_file(path: str | Path, baseline: str | Path) -> dict[str, Any]:
    """Diff `baseline` (before) against `path` (after)."""
    kind = kind_of(path)
    if kind == "docx":
        return diff_docx(path, baseline)
    if kind == "pptx":
        return diff_pptx(baseline, path, detail="text")
    return diff_xlsx(baseline, path)
