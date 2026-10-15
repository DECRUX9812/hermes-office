---
name: office-review
description: Review office docs with tracked changes.
version: 0.1.0
author: Factory Droid
license: MIT
platforms: [linux, macos, windows]
tags: [office, docx, pptx, xlsx]
category: productivity
metadata:
  hermes:
    related_skills: []
---

# Office Review Skill

Use this skill when a document needs inspection, search, a tracked edit, a comment, a redline decision, or a receipt lookup. It does not rewrite files silently, and it refuses to guess when a target is ambiguous.

## When to Use

- Inspect a Word, PowerPoint, or Excel file before changing it.
- Search for text before choosing a target.
- Make a tracked docx edit or comment.
- Accept or reject tracked Word changes only when asked.
- Revert a receipt after review.

## Procedure

1. Inspect with `office_inspect`.
2. Search with `office_search` if the target is unclear.
3. Propose the smallest change.
4. Use `docx_tracked_edit`, `docx_comment`, `pptx_edit_shape`, or `xlsx_edit`.
5. Read the receipt with `office_receipt`.
6. Mark the workstream `needs_you` when the human should review it.

## Pitfalls

- If a docx target is ambiguous, narrow it with `nth`.
- If a PowerPoint shape is missing, refind the slide or shape name.
- If a workbook formula depends on moved cells, recalc after the edit.
- If the tool refuses, do not try a different hack. Narrow the target or ask.

## Verification

- Inspect the diff.
- Confirm the receipt.
- Confirm the workstream status.
