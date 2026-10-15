from __future__ import annotations

from pathlib import Path
from typing import Any

from pptx import Presentation
from pptx.edit import refind, replace_text, replace_text_at
from pptx.errors import AmbiguousTargetError, TargetNotFoundError
from pptx.inspect import effective_font, effective_paragraph_format, effective_shape_format, inspect_deck, inspect_text
from pptx.package import patch_save
from pptx.diff import diff_decks

from .guard import resolve_path


def _load(path: str | Path):
    return Presentation(str(resolve_path(path)))


def _save(path: str | Path, prs) -> Path:
    target = resolve_path(path)
    if target.exists():
        tmp = target.parent / f".{target.name}.patched.pptx"
        patch_save(str(target), prs, str(tmp))
        tmp.replace(target)
    else:
        prs.save(str(target))
    return target


def inspect_pptx(path: str | Path) -> dict[str, Any]:
    prs = _load(path)
    deck = inspect_deck(prs).to_dict()
    return {"kind": "pptx", "path": str(resolve_path(path)), "deck": deck}


def inspect_pptx_formats(path: str | Path, *, slide: int, shape: str) -> dict[str, Any]:
    prs = _load(path)
    shp = prs.slides[slide].shapes.shape_by_name(shape)
    paragraph = shp.text_frame.paragraphs[0] if getattr(shp, "has_text_frame", False) and shp.text_frame.paragraphs else None
    run = paragraph.runs[0] if paragraph and paragraph.runs else None
    return {
        "kind": "pptx",
        "path": str(resolve_path(path)),
        "font": effective_font(run).to_dict() if run else None,
        "paragraph": effective_paragraph_format(paragraph).to_dict() if paragraph else None,
        "shape": effective_shape_format(shp).to_dict(),
    }


def edit_pptx_shape(
    path: str | Path,
    *,
    slide: int,
    shape: str,
    ops: list[dict[str, Any]],
) -> dict[str, Any]:
    prs = _load(path)
    current_slide = prs.slides[slide]
    shp = current_slide.shapes.shape_by_name(shape)
    slide_ops: list[dict[str, Any]] = []
    for op in ops:
        kind = op.get("kind") or op.get("op")
        if kind == "replace_text":
            replace_text_at(prs, refind(prs, op["anchor"]), op["find"], op["replace"])
        elif kind == "set_geometry":
            shp.left = int(op["x"])
            shp.top = int(op["y"])
            shp.width = int(op["w"])
            shp.height = int(op["h"])
        elif kind == "delete":
            current_slide.shapes.delete(shp)
        elif kind == "clone_slide":
            clone = prs.slides.clone(current_slide, after=op.get("after"))
            slide_ops.append({"clone": clone.slide_id})
        elif kind == "delete_slide":
            prs.slides.delete(current_slide)
            slide_ops.append({"delete_slide": slide})
        elif kind == "move_slide":
            prs.slides.move(current_slide, int(op["to_index"]))
            slide_ops.append({"move_slide": int(op["to_index"])})
        elif kind == "replace_text_shape":
            replace_text(prs, op["find"], op["replace"], include_notes=bool(op.get("include_notes", False)))
        else:
            raise ValueError(f"unknown pptx op: {kind}")
    _save(path, prs)
    return {"ok": True, "slide_ops": slide_ops}


def compose_pptx_decks(dest: str | Path, sources: list[dict[str, Any]], *, mode: str, output_path: str | Path) -> dict[str, Any]:
    out = Path(output_path).expanduser().resolve()
    prs = _load(dest) if Path(dest).exists() else Presentation()
    for source in sources:
        src_prs = _load(source["path"])
        slide_indexes = source.get("slides")
        if slide_indexes is None:
            prs.append_deck(src_prs, mode=mode, notes=True)
            continue
        for slide_index in slide_indexes:
            prs.import_slide(src_prs, src_prs.slides[slide_index], mode=mode, notes=True)
    _save(out, prs)
    return {"ok": True, "output_path": str(out)}


def diff_pptx(path_a: str | Path, path_b: str | Path, *, detail: str = "text") -> dict[str, Any]:
    return diff_decks(path_a, path_b, detail=detail).to_dict()


def preview_pptx(path: str | Path) -> dict[str, Any]:
    prs = _load(path)
    slides = []
    for index, slide in enumerate(prs.slides):
        shapes = []
        for shape in slide.shapes:
            shapes.append(
                {
                    "name": shape.name,
                    "kind": str(getattr(shape, "shape_type", None)),
                    "text": getattr(shape, "text", None) if getattr(shape, "has_text_frame", False) else None,
                    "x": getattr(shape, "left", None),
                    "y": getattr(shape, "top", None),
                    "w": getattr(shape, "width", None),
                    "h": getattr(shape, "height", None),
                }
            )
        slides.append({"index": index, "title": getattr(slide.shapes.title, "text", None), "shapes": shapes})
    return {"kind": "pptx", "path": str(resolve_path(path)), "pptx": {"slides": slides, "slide_w": prs.slide_width, "slide_h": prs.slide_height}}
