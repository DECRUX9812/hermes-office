---
name: office-review
description: Review and edit Word, PowerPoint and Excel files on the user's Desk with tracked changes, backups and receipts. Use for any .docx/.pptx/.xlsx inspection, edit, redline, recalc or diff.
version: 0.1.0
license: MIT
tags: [office, docx, pptx, xlsx]
---

# Office work on the Desk

The Desk is the set of folders the user has trusted you to edit (`office_desk action=list`).
Writes outside a Desk folder are refused with `outside_desk`; reads work anywhere.

Every edit is backed up first, written atomically, and receipted (before/after sha256, backup
path, the library's own report). Word edits are tracked changes the user can accept or reject.
Never work around a refusal: narrow the target or ask.

## Procedure

1. `office_inspect` the file first (detail=full when you need comment ids or revision indexes).
2. If the target text might occur more than once, `office_search` first, then pass `nth`.
3. Make the smallest edit that satisfies the request.
4. Read the returned `receipt_id`; the change shows up on the Desk page automatically.
5. End your turn with `office_desk action=needs_you note="what to review"` when the user should
   look at something, or `action=done` when finished.

## Tool notes

- `docx_tracked_edit` refuses ambiguous (`AmbiguousTargetError` → add `nth`) and missing
  (`TargetNotFoundError`) targets. Nothing is written on a refusal.
- Accept/reject tracked changes ONLY when the user asked: `docx_redline_accept_reject`.
  `action=compare` with `compare_with` builds a redline (tracked diff between two versions).
- `pptx_edit_shape` addresses shapes by name and slides by 0-based index; get both from
  `office_inspect`. `mode` for `pptx_compose_decks` is required: `adopt_theme`,
  `keep_appearance` or `bake`.
- `xlsx_edit` preserves styles/charts/pivots; formulas are rewritten when rows/cols move.
  After edits that formulas depend on, run `xlsx_recalc` (uses LibreOffice; reports values).
- Undo anything with `office_receipt action=revert receipt_id=...`.

## Pitfalls

- Only .docx/.pptx/.xlsx (legacy .doc/.ppt/.xls and macro-enabled files are refused).
- The same chat session reuses one workstream per folder; pass `workstream_id` when the Desk
  page created one for this request.
- If `xlsx_recalc` reports `libreoffice_missing`, tell the user instead of editing cached values.
