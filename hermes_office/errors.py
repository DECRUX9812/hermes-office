from __future__ import annotations


class OfficeError(Exception):
    reason = "office_error"
    hint = "Check the request and try again."


class PolicyRefusal(OfficeError):
    """A desk policy said no before the file was touched."""

    reason = "policy_refusal"


class OutsideDeskError(PolicyRefusal):
    reason = "outside_desk"
    hint = "Only files inside a Desk folder can be edited. Ask the user, then office_desk action=add_folder path=<folder>."


class UnsupportedFileError(PolicyRefusal):
    reason = "unsupported_file"
    hint = "Use a .docx, .pptx or .xlsx file."


class NotFoundError(OfficeError):
    reason = "not_found"
    hint = "Check the id or path; office_desk action=list shows folders and workstreams."
