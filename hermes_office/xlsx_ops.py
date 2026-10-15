from __future__ import annotations

import os
import shutil
import tempfile
from pathlib import Path
from typing import Any

from openpyxl import load_workbook, oracle
from openpyxl.errors import TargetNotFoundError
from openpyxl.preserve import diff_workbooks, scan_errors

from . import settings
from .guard import resolve_path

PREVIEW_ROWS = 60
PREVIEW_COLS = 20


def _load(path: str | Path, *, data_only: bool = False):
    return load_workbook(str(resolve_path(path)), data_only=data_only)


def _sheet(wb, name: str | None):
    if name is None:
        return wb.worksheets[0]
    if name not in wb.sheetnames:
        raise TargetNotFoundError(f"sheet {name!r} not found (sheets: {wb.sheetnames}); nothing was changed")
    return wb[name]


def _split_ref(ref: str, default_sheet: str | None) -> tuple[str | None, str]:
    if "!" in ref:
        sheet, cell = ref.rsplit("!", 1)
        return sheet.strip("'"), cell
    return default_sheet, ref


def _is_formula(value: Any) -> bool:
    return isinstance(value, str) and value.startswith("=")


def inspect_xlsx(path: str | Path, *, detail: str = "summary") -> dict[str, Any]:
    wb = _load(path)
    sheets = []
    for ws in wb.worksheets:
        formulas = [cell.coordinate for row in ws.iter_rows() for cell in row if _is_formula(cell.value)]
        entry: dict[str, Any] = {"name": ws.title, "dims": ws.calculate_dimension(), "max_row": ws.max_row, "max_col": ws.max_column, "formula_cells": len(formulas)}
        sample = formulas if detail == "full" else formulas[:10]
        entry["formulas"] = [{"ref": ref, "formula": ws[ref].value} for ref in sample]
        sheets.append(entry)
    return {
        "kind": "xlsx",
        "path": str(resolve_path(path)),
        "summary": {"sheets": len(sheets), "defined_names": list(wb.defined_names.keys())},
        "sheets": sheets,
        "errors": scan_errors(wb),
    }


def search_xlsx(path: str | Path, query: str, *, regex: bool = False) -> list[dict[str, Any]]:
    return list(_load(path).search(query, regex=regex))


def edit_xlsx(path: str | Path, *, ops: list[dict[str, Any]]) -> dict[str, Any]:
    wb = _load(path)
    applied: list[dict[str, Any]] = []
    remaps = []
    for op in ops:
        kind = op.get("op")
        if kind in {"set", "formula"}:
            sheet_name, cell = _split_ref(op["ref"], op.get("sheet"))
            ws = _sheet(wb, sheet_name)
            value = op["number"] if kind == "set" and op.get("number") is not None else op["value"]
            if kind == "formula" and not _is_formula(value):
                value = f"={value}"
            before = ws[cell].value
            ws[cell] = value
            applied.append({"op": kind, "ref": f"{ws.title}!{cell}", "before": before, "after": value})
        elif kind in {"insert_rows", "delete_rows", "insert_cols", "delete_cols"}:
            ws = _sheet(wb, op.get("sheet"))
            remap = getattr(ws, kind)(int(op["index"]), int(op.get("amount", 1)))
            if remap is not None:
                remaps.append(remap)
            applied.append({"op": kind, "sheet": ws.title, "index": int(op["index"]), "amount": int(op.get("amount", 1))})
        elif kind == "rename_sheet":
            ws = _sheet(wb, op.get("sheet"))
            old = ws.title
            ws.title = op["new_name"]
            applied.append({"op": kind, "from": old, "to": ws.title})
        else:
            raise ValueError(f"unknown xlsx op {kind!r}; use set, formula, insert_rows, delete_rows, insert_cols, delete_cols, rename_sheet")
    # Preserve-mode save delivers atomically and reports exactly what changed against the loaded bytes.
    receipt = wb.save(str(resolve_path(path)), receipt=True)
    return {"applied": applied, "edit_receipt": receipt.to_dict(), "remaps": [repr(remap) for remap in remaps]}


def _read_values(path: str | Path, refs: list[str]) -> dict[str, Any]:
    wb = _load(path, data_only=True)
    values = {}
    for ref in refs:
        sheet_name, cell = _split_ref(ref, None)
        values[ref] = _sheet(wb, sheet_name)[cell].value
    return values


def soffice_path() -> str | None:
    return oracle.find_soffice()


def recalc_xlsx(path: str | Path, *, read: list[str] | None = None, timeout: float | None = None) -> dict[str, Any]:
    """Recalculate through LibreOffice into a candidate, then swap it in. Never claims success without LibreOffice."""
    target = resolve_path(path)
    read = list(read or [])
    if soffice_path() is None:
        return {
            "recalculated": False,
            "reason": "libreoffice_missing",
            "hint": "Install LibreOffice (soffice) to recalculate; cached values are unchanged.",
            "errors": scan_errors(_load(target)),
            "values": _read_values(target, read),
        }
    workdir = Path(tempfile.mkdtemp(prefix=".hermes-office-recalc-", dir=str(target.parent)))
    try:
        candidate = workdir / target.name
        result = oracle.recalc(str(target), output_path=str(candidate), timeout=timeout if timeout is not None else settings.recalc_timeout())
        os.replace(candidate, target)
    finally:
        shutil.rmtree(workdir, ignore_errors=True)
    return {"recalculated": True, "status": result.status, "result": result.to_dict(), "values": _read_values(target, read)}


def diff_xlsx(path_a: str | Path, path_b: str | Path) -> dict[str, Any]:
    report = diff_workbooks(str(resolve_path(path_a)), str(resolve_path(path_b)))
    return {"kind": "xlsx", **report.to_dict()}


def preview_xlsx(path: str | Path) -> dict[str, Any]:
    wb = _load(path)
    cached = _load(path, data_only=True)
    sheets = []
    for ws in wb.worksheets:
        values_ws = cached[ws.title]
        max_row = min(PREVIEW_ROWS, ws.max_row or 1)
        max_col = min(PREVIEW_COLS, ws.max_column or 1)
        rows = []
        for row in ws.iter_rows(min_row=1, max_row=max_row, min_col=1, max_col=max_col):
            cells = []
            for cell in row:
                formula = cell.value if _is_formula(cell.value) else None
                value = values_ws[cell.coordinate].value if formula else cell.value
                cells.append({"ref": cell.coordinate, "value": value, "formula": formula})
            rows.append(cells)
        sheets.append({"name": ws.title, "dims": ws.calculate_dimension(), "rows": rows, "truncated": (ws.max_row or 0) > max_row or (ws.max_column or 0) > max_col})
    return {"kind": "xlsx", "path": str(resolve_path(path)), "xlsx": {"sheets": sheets}}
