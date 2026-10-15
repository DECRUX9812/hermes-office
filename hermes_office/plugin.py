from __future__ import annotations

from pathlib import Path

from . import bridge
from .schemas import (
    DOCX_COMMENT_SCHEMA,
    DOCX_REDLINE_SCHEMA,
    DOCX_TRACKED_EDIT_SCHEMA,
    OFFICE_DESK_SCHEMA,
    OFFICE_INSPECT_SCHEMA,
    OFFICE_RECEIPT_SCHEMA,
    OFFICE_SEARCH_SCHEMA,
    PPTX_COMPOSE_SCHEMA,
    PPTX_DIFF_SCHEMA,
    PPTX_EDIT_SHAPE_SCHEMA,
    PPTX_INSPECT_FORMATS_SCHEMA,
    XLSX_DIFF_SCHEMA,
    XLSX_EDIT_SCHEMA,
    XLSX_RECALC_SCHEMA,
)
from .tools import (
    docx_comment,
    docx_redline_accept_reject,
    docx_tracked_edit,
    office_desk,
    office_inspect,
    office_receipt_handler,
    office_search,
    pptx_compose_decks_handler,
    pptx_diff_handler,
    pptx_edit_shape_handler,
    pptx_inspect_formats_handler,
    xlsx_diff_handler,
    xlsx_edit_handler,
    xlsx_recalc_handler,
)


def _check_requirements() -> bool:
    try:
        import docx
        import pptx
        import openpyxl

        return bool(getattr(docx, "__paper_version__", None) and getattr(pptx, "__paper_version__", None) and getattr(openpyxl, "__paper_version__", None))
    except Exception:
        return False


def register(ctx) -> None:
    skill = Path(__file__).parent / "skills" / "office-review" / "SKILL.md"
    tools = [
        ("office_desk", OFFICE_DESK_SCHEMA, office_desk),
        ("office_inspect", OFFICE_INSPECT_SCHEMA, office_inspect),
        ("office_search", OFFICE_SEARCH_SCHEMA, office_search),
        ("docx_tracked_edit", DOCX_TRACKED_EDIT_SCHEMA, docx_tracked_edit),
        ("docx_comment", DOCX_COMMENT_SCHEMA, docx_comment),
        ("docx_redline_accept_reject", DOCX_REDLINE_SCHEMA, docx_redline_accept_reject),
        ("pptx_inspect_formats", PPTX_INSPECT_FORMATS_SCHEMA, pptx_inspect_formats_handler),
        ("pptx_edit_shape", PPTX_EDIT_SHAPE_SCHEMA, pptx_edit_shape_handler),
        ("pptx_compose_decks", PPTX_COMPOSE_SCHEMA, pptx_compose_decks_handler),
        ("pptx_diff", PPTX_DIFF_SCHEMA, pptx_diff_handler),
        ("xlsx_edit", XLSX_EDIT_SCHEMA, xlsx_edit_handler),
        ("xlsx_recalc", XLSX_RECALC_SCHEMA, xlsx_recalc_handler),
        ("xlsx_diff", XLSX_DIFF_SCHEMA, xlsx_diff_handler),
        ("office_receipt", OFFICE_RECEIPT_SCHEMA, office_receipt_handler),
    ]
    for name, schema, handler in tools:
        ctx.register_tool(name=name, toolset="office", schema=schema, handler=handler, check_fn=_check_requirements, emoji="📎")
    ctx.register_skill("office-review", skill)
