# hermes-office

Hermes plugin for office docs on a desk, using paper-docx, paper-pptx, and paper-xlsx.

## Install

```bash
pip install .
```

## Tools

| Tool | Purpose |
|---|---|
| `office_desk` | manage desk folders and workstreams |
| `office_inspect` | summarize a docx, pptx, or xlsx file |
| `office_search` | search office text |
| `docx_tracked_edit` | tracked docx text replacement |
| `docx_comment` | add or reply to comments |
| `docx_redline_accept_reject` | accept or reject tracked changes |
| `pptx_inspect_formats` | inspect slide formatting |
| `pptx_edit_shape` | edit or remove a named shape |
| `pptx_compose_decks` | import slides between decks |
| `pptx_diff` | diff two decks |
| `xlsx_edit` | edit workbook cells or structure |
| `xlsx_recalc` | recalculate formulas |
| `xlsx_diff` | diff two workbooks |
| `office_receipt` | list or revert receipts |

## Safety

Writes are desk-guarded. Outside-desk edits refuse instead of touching files.
